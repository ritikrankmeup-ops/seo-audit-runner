#!/usr/bin/env python3
"""
Build the SEO audit workbook (.xlsx, one tab per area) from crawl.json (+ optional analysis.json).

  python3 build_report.py crawl.json --analysis analysis.json --out report.xlsx [--csv-dir csv/]
  python3 build_report.py crawl.json --summary            # condensed findings for Claude to read

analysis.json (written by Claude, all keys optional):
{
  "audit_label": "Full Website Audit",
  "scores": {"Technical SEO": 60, "On-Page SEO": 55, ...},
  "executive_summary": "3-5 sentences",
  "issues": [{"area": "On-Page", "issue": "...", "severity": "High", "evidence": "...",
              "urls": ["..."], "why": "...", "fix": "...", "owner": "Content team"}],
  "page_notes": {"https://url": "short judgement for this page"},
  "action_plan": [{"phase": "Week 1", "task": "...", "owner": ""}],
  "offpage": {"backlinks_note": "...", "referring_domains": [{"domain": "...", "source": "...", "note": ""}],
              "brand_mentions": [], "notes": ["..."]}
}
"""
import argparse
import csv
import difflib
import json
import os
import re
from collections import Counter, defaultdict
from urllib.parse import urlparse

SEV_ORDER = ["Critical", "High", "Medium", "Low", "Info"]
SEV_FILL = {"Critical": "C00000", "High": "F4B084", "Medium": "FFE699", "Low": "C6E0B4", "Info": "DDEBF7"}
LEGACY_IMG = (".jpg", ".jpeg", ".png", ".gif", ".bmp")
EXPECTED_SCHEMA = {
    "Homepage": ["Organization", "WebSite"],
    "Product page": ["Product", "BreadcrumbList"],
    "Product category / shop": ["CollectionPage|ItemList", "BreadcrumbList"],
    "Blog post / article": ["Article|BlogPosting|NewsArticle", "BreadcrumbList"],
    "Blog archive / category": ["CollectionPage|Blog", "BreadcrumbList"],
    "Location page": ["LocalBusiness", "BreadcrumbList"],
    "Contact page": ["ContactPage|LocalBusiness|Organization"],
    "About page": ["AboutPage|Organization"],
    "Landing / service page": ["Service|WebPage", "BreadcrumbList"],
}
LOCAL_TYPES = {"localbusiness", "dentist", "medicalclinic", "medicalbusiness", "physician", "store",
               "restaurant", "professionalservice", "homeandconstructionbusiness", "legalservice",
               "automotivebusiness", "healthandbeautybusiness", "hairsalon", "beautysalon",
               "dayspa", "medicalorganization", "hospital", "lodgingbusiness", "hotel",
               "realestateagent", "financialservice", "accountingservice", "emergencyservice",
               "autorepair", "plumber", "electrician", "generalcontractor", "lawyer", "attorney", "optician"}
# Any of these satisfies "Organization" (schema.org subtypes of Organization, incl. every LocalBusiness)
ORG_TYPES = {"organization", "corporation", "medicalorganization", "educationalorganization",
             "collegeoruniversity", "school", "highschool", "middleschool", "elementaryschool", "preschool",
             "ngo", "onlinestore", "onlinebusiness", "newsmediaorganization", "sportsorganization",
             "governmentorganization", "researchorganization", "airline", "library",
             "performinggroup", "consortium"} | LOCAL_TYPES
SOCIAL = ("facebook.com", "instagram.com", "linkedin.com", "twitter.com", "x.com", "youtube.com",
          "pinterest.com", "tiktok.com", "wa.me", "whatsapp.com", "threads.net")


# ============================================================================ where-to-fix guide
# Matched by issue-title prefix. Keys: platform family from the crawler, "default" as fallback.
WP_SEO = "Rank Math / Yoast"
GUIDE = [
    ("Internal URLs that redirect", {
        "wordpress": "Edit the source page (Elementor/Gutenberg) or Appearance > Menus and replace the link. Bulk: Tools > Better Search Replace (old URL -> final URL).",
        "shopify": "Online Store > Navigation (menus) or the page/product editor: replace the link.",
        "nextjs": "Search the codebase/CMS for the old path and change the href in the component or content entry.",
        "react": "Search the codebase/CMS for the old path and change the href in the component or content entry.",
        "default": "Edit the link in the page or template source."}),
    ("Redirect chains", {
        "wordpress": f"{WP_SEO} > Redirections (or the Redirection plugin): point the FIRST rule straight to the final URL; delete the in-between rules.",
        "nextjs": "next.config.js > redirects(): set the destination of the first rule to the final URL.",
        "default": "Web server / CDN redirect rules: one 301 from the first URL to the final URL."}),
    ("Redirecting URLs listed in the XML sitemap", {
        "wordpress": f"The sitemap is built by {WP_SEO}. Update the old post/page URL, or trash/noindex the old page so it leaves the sitemap.",
        "nextjs": "app/sitemap.ts (or next-sitemap config): output the final URL instead of the old one.",
        "default": "Sitemap generator / sitemap file: replace each URL with its final URL (see 'Fix List')."}),
    ("Broken (4xx/5xx) URLs listed in the XML sitemap", {
        "wordpress": f"Restore the page, or trash it so {WP_SEO} drops it from the sitemap; add a 301 if it had traffic.",
        "nextjs": "app/sitemap.ts / next-sitemap: remove these paths (they no longer have a route).",
        "default": "Remove these URLs from the sitemap file/generator."}),
    ("Broken internal URLs", {
        "wordpress": f"Edit each source page listed in 'Fix List' and change/remove the link, or add a 301 in {WP_SEO} > Redirections to the suggested page.",
        "nextjs": "Change the href in the component/content, or add a redirect in next.config.js > redirects().",
        "default": "Fix the link on each source page (see 'Fix List'), or 301 the dead URL to the suggested page."}),
    ("Pages linking to redirecting URLs", {"default": "Edit the link on each page listed in 'Fix List' to the final URL."}),
    ("Duplicate URLs across sitemaps", {"wordpress": f"{WP_SEO} > Sitemap settings: disable duplicate sitemaps (e.g. a second SEO/sitemap plugin).",
                                        "default": "Sitemap index: make child sitemaps non-overlapping."}),
    ("HTTP version does not redirect", {
        "wordpress": "Settings > General: both URLs https://. Then force HTTPS at host (cPanel > Domains > Force HTTPS) or Cloudflare > SSL/TLS > Edge > Always Use HTTPS.",
        "default": "Web server / CDN: 301 all http:// requests to https://."}),
    ("Both www and non-www", {"default": "Web server / CDN (Cloudflare Redirect Rule): 301 the non-preferred host to the preferred one."}),
    ("Non-existent URLs return 200", {
        "wordpress": "Disable any '404 to homepage' plugin (e.g. All 404 Redirect to Homepage); check the theme's 404.php sends status 404.",
        "nextjs": "In dynamic routes call notFound() when no data is found (App Router) or return { notFound: true } (Pages Router).",
        "react": "A client-side SPA always returns 200. Pre-render/SSR the routes, or add a server/CDN rule that returns 404 for unknown paths.",
        "default": "Make the app/server send HTTP status 404 for unknown URLs (not a 200 'not found' page)."}),
    ("robots.txt missing", {"wordpress": f"{WP_SEO} > General Settings > Edit robots.txt.", "default": "Add /robots.txt at the site root."}),
    ("No XML sitemap", {"wordpress": f"{WP_SEO} > Sitemap: enable, then submit in Search Console > Sitemaps.",
                        "nextjs": "Add app/sitemap.ts (or next-sitemap) and submit it in Search Console.",
                        "default": "Generate /sitemap.xml and submit it in Search Console."}),
    ("Noindex pages listed in the sitemap", {"wordpress": "Edit page > Rank Math/Yoast > Advanced > Robots: set Index if it should rank; otherwise it leaves the sitemap automatically.",
                                             "default": "Either remove noindex or remove the URL from the sitemap."}),
    ("Missing canonical", {"wordpress": "Rank Math/Yoast add self-canonicals automatically: check no plugin/theme disables them.",
                           "nextjs": "In the route: export const metadata = { alternates: { canonical: '<path>' } } (or generateMetadata).",
                           "default": "Add <link rel=\"canonical\" href=\"<page URL>\"> in <head>."}),
    ("Canonical points to a redirecting", {"wordpress": "Edit page > Rank Math/Yoast > Advanced > Canonical URL: set the final URL (or clear to self).",
                                           "default": "Set the canonical to the final 200 URL."}),
    ("Canonical points to a different", {"default": "Review in 'Tech - Indexability'. If the page should rank, make the canonical self-referencing."}),
    ("Slow server response", {"wordpress": "Enable page caching: LiteSpeed Cache / WP Rocket; Cloudflare APO; check PHP version and hosting.",
                              "nextjs": "Use static generation / ISR (export const revalidate = 3600) for public pages; cache at CDN (s-maxage).",
                              "default": "Add full-page caching at server/CDN; reduce server work per request."}),
    ("Mixed content", {"wordpress": "Tools > Better Search Replace: http://yourdomain -> https://yourdomain (all tables).",
                       "default": "Change every http:// resource URL to https:// in templates/content."}),
    ("Missing viewport", {"wordpress": "Theme header.php (or a child theme) <head>.", "nextjs": "app/layout: export const viewport = { width: 'device-width', initialScale: 1 }.",
                          "default": "Add the viewport meta tag in the base HTML template <head>."}),
    ("Missing <html lang>", {"wordpress": "Settings > General > Site Language.", "nextjs": "app/layout.tsx: <html lang=\"en\">.",
                             "default": "Base HTML template: <html lang=\"en\">."}),
    ("Same URL linked with and without", {"default": "Pick one URL format; fix internal links and 301 the other form."}),
    ("Uppercase letters", {"default": "Use lowercase slugs; 301 uppercase versions to lowercase."}),
    ("Orphan pages", {"default": "Add a contextual link from the hub page suggested in 'Fix List' (category/parent page), or remove obsolete pages."}),
    ("Poor mobile performance", {"wordpress": "Optimise images (ShortPixel/Imagify + WebP), delay JS (WP Rocket/LiteSpeed), remove unused plugins, page cache.",
                                 "nextjs": "Use next/image, dynamic import heavy components, reduce client JS, cache pages (ISR).",
                                 "default": "Fix the 'Top fixes' listed per URL in 'Tech - Core Web Vitals'."}),
    ("Real-user", {"default": "See 'Tech - Core Web Vitals' (top fixes per URL); re-check after 28 days in Search Console > Core Web Vitals."}),
    ("Core Web Vitals not measured", {"default": "claude.ai/code > environment settings > Environment variables: PAGESPEED_API_KEY=<free Google API key>."}),
    ("Missing or empty <title>", {"wordpress": "Edit page > Rank Math/Yoast snippet editor > Title.",
                                  "nextjs": "Route file: export const metadata = { title: '...' } or generateMetadata().",
                                  "default": "Set <title> in the page template/CMS field."}),
    ("Title ", {"wordpress": "Edit page > Rank Math/Yoast snippet editor > Title. Templated titles: Rank Math > Titles & Meta > [post type] > Single Title format.",
                "shopify": "Product/page editor > Search engine listing > Page title.",
                "nextjs": "generateMetadata()/metadata.title in the route; site-wide suffix: title.template in app/layout.",
                "default": "Title field in the CMS / page template (see suggested titles in 'Fix List')."}),
    ("Duplicate titles", {"wordpress": "Edit each page > Rank Math/Yoast > Title. If many share one title, fix the title template.",
                          "nextjs": "generateMetadata(): build the title from the page's own data (not a shared constant).",
                          "default": "Give each URL its own title (see 'Fix List')."}),
    ("Missing meta description", {"wordpress": "Edit page > Rank Math/Yoast snippet editor > Description.",
                                  "nextjs": "metadata.description / generateMetadata() per route.",
                                  "default": "Meta description field in the CMS / page template."}),
    ("Meta description", {"wordpress": "Edit page > Rank Math/Yoast snippet editor > Description.",
                          "nextjs": "metadata.description / generateMetadata() per route.",
                          "default": "Meta description field in the CMS / template (see 'Fix List')."}),
    ("Duplicate meta descriptions", {"default": "Write a unique description per URL (see 'Fix List')."}),
    ("Missing H1", {"wordpress": "Elementor: Heading widget > HTML Tag = H1 (one per page). Gutenberg: Heading block > H1.",
                    "nextjs": "Page component: wrap the main page heading in <h1>.", "default": "Mark the main heading as <h1> in the page template."}),
    ("Multiple H1", {"wordpress": "Elementor: select the extra heading > Content > HTML Tag = H2. Gutenberg: change block level to H2.",
                     "nextjs": "Page component: keep one <h1>, change the others to <h2>.", "default": "Keep one <h1>; change the others to <h2>."}),
    ("H1 inside header", {"wordpress": "Elementor > Templates > Theme Builder > Header: logo/site title widget HTML Tag = div.",
                          "default": "Header template/component: change the <h1> to <div> or <p>."}),
    ("Heading hierarchy skips", {"wordpress": "Elementor Heading widget > HTML Tag (choose the next level down).", "default": "Change the heading level in the page source."}),
    ("Content starts with", {"default": "Move the H1 above other headings in the page layout."}),
    ("Empty heading", {"default": "Delete the empty <hN> element in the template."}),
    ("Headings used in header", {
        "wordpress": "Elementor > Templates > Theme Builder > Header/Footer > click the heading > Content > HTML Tag = div (style stays). Widget titles: Appearance > Widgets.",
        "nextjs": "Header/Footer component: replace <hN> with <p className=\"...\"> (keep the same class).",
        "react": "Header/Footer component: replace <hN> with <p className=\"...\"> (keep the same class).",
        "default": "Header/footer template: replace the heading tag with <p>/<div>, keeping the CSS class."}),
    ("Thin content", {"default": "Add useful content to the page: FAQs, process, benefits, proof (reviews, case studies), local details."}),
    ("Missing Open Graph", {"wordpress": "Edit page > Rank Math/Yoast > Social tab: title + 1200x630 image.",
                            "nextjs": "metadata.openGraph = { title, images: [...] } in the route/layout.", "default": "Add og:title and og:image meta tags."}),
    ("Images larger than 100 KB", {"wordpress": "Install ShortPixel / Imagify / EWWW > Bulk optimise + WebP delivery. Re-upload hero images at display width (≤1920px, ≤150 KB).",
                                   "shopify": "Re-upload compressed images (Shopify serves WebP automatically via image_url filters).",
                                   "nextjs": "Use next/image (automatic WebP/AVIF + sizes); compress originals in /public or the CMS.",
                                   "default": "Compress, resize to display size and serve WebP/AVIF."}),
    ("Large images in legacy format", {"wordpress": "ShortPixel/Imagify 'Create WebP' + 'Deliver WebP' option.", "default": "Convert to WebP/AVIF."}),
    ("Content images missing alt", {"wordpress": "Media > Library > click image > Alternative Text (or Elementor Image widget > Alt).",
                                    "nextjs": "Add alt=\"...\" to each <Image>/<img> (or the CMS alt field).", "default": "Add alt text in the CMS media field or <img alt>."}),
    ("Broken images", {"default": "Re-upload the image or remove the <img> from the page."}),
    ("Images without width", {"wordpress": "Use the image widget's 'Image size' setting; the theme should output width/height.",
                              "nextjs": "next/image requires width/height (or fill) - use it.", "default": "Add width and height attributes to <img>."}),
    ("Invalid JSON-LD", {"default": "Fix the JSON syntax in the <script type=\"application/ld+json\"> block; test at validator.schema.org."}),
    ("Pages with no structured data", {"wordpress": "Rank Math > Titles & Meta / Schema (or 'Schema & Structured Data for WP' plugin).",
                                       "nextjs": "Render a <script type=\"application/ld+json\"> with JSON.stringify(schema) in the page/layout.",
                                       "default": "Add JSON-LD in the page <head> (snippets in 'Fix List')."}),
    ("Schema type expected", {"wordpress": "Rank Math > Schema tab on the page (or WPCode > Header snippet for custom JSON-LD).",
                              "nextjs": "Add the JSON-LD from 'Fix List' to that route's page component.",
                              "default": "Paste the JSON-LD from 'Fix List' into the page <head>."}),
    ("No Organization schema", {"wordpress": "Rank Math > Titles & Meta > Local SEO / Yoast > Settings > Site representation: Organization, logo, social profiles.",
                                "nextjs": "app/layout: render the Organization JSON-LD (snippet in 'Fix List') once, site-wide.",
                                "default": "Add the Organization JSON-LD (in 'Fix List') to the site-wide header."}),
    ("Organization name differs", {"default": "Output ONE Organization node with a fixed \"@id\": \"https://<domain>/#organization\"; remove duplicates from other plugins/templates."}),
    ("Review/rating markup", {"wordpress": "Find which plugin/theme adds aggregateRating (view source, search 'aggregateRating') and disable that option.",
                              "default": "Remove the aggregateRating/review property from the Organization/LocalBusiness JSON-LD template."}),
    ("Broken external links", {"default": "Edit the source pages listed in 'OffPage - Links & Profiles' and update/remove the link."}),
    ("Social profiles linked", {"default": "Add the profile URLs to Organization \"sameAs\" (snippet in 'Fix List')."}),
    ("Phone number on site differs", {"default": "Use one phone number/format in header, footer, contact page, schema and Google Business Profile."}),
]

SCOPES = {  # which issues/tabs each audit type returns
    "seo-technical": {"areas": {"Technical"}, "categories": {"Invalid JSON-LD"},
                      "tabs": ["Fix List", "Issue Tracker", "Page Inventory", "Page Types", "Tech - Redirects",
                               "Tech - URL Errors", "Tech - Indexability", "Tech - Sitemap URLs", "Tech - Site Checks",
                               "Tech - Core Web Vitals", "Action Plan", "Data Sources"]},
    "seo-sitemap": {"areas": {"Technical"}, "only_categories": {"Sitemap", "Redirects", "URL errors", "Indexability", "Robots"},
                    "tabs": ["Fix List", "Issue Tracker", "Tech - Sitemap URLs", "Tech - Redirects", "Tech - URL Errors",
                             "Tech - Indexability", "Action Plan", "Data Sources"]},
    "seo-schema": {"areas": {"Schema"}, "tabs": ["Fix List", "Issue Tracker", "Schema - Per Page", "Schema - Org & Local",
                                                 "Page Types", "Action Plan", "Data Sources"]},
    "seo-images": {"areas": {"On-Page"}, "only_categories": {"Images"},
                   "tabs": ["Fix List", "Issue Tracker", "OnPage - Images", "Tech - Core Web Vitals", "Action Plan", "Data Sources"]},
    "seo-content": {"areas": {"On-Page", "Content"}, "exclude_categories": {"Images"},
                    "tabs": ["Fix List", "Issue Tracker", "Page Inventory", "OnPage - Titles & Meta", "OnPage - Headings",
                             "OnPage - Header-Footer H", "Page Types", "Action Plan", "Data Sources"]},
    "seo-local": {"areas": {"Schema", "Off-Page", "Local"},
                  "tabs": ["Fix List", "Issue Tracker", "Schema - Org & Local", "OffPage - Links & Profiles", "Action Plan", "Data Sources"]},
    "seo-geo": {"areas": {"Schema", "Content", "AI Search", "On-Page"}, "exclude_categories": {"Images", "Meta description"},
                "tabs": ["Fix List", "Issue Tracker", "Schema - Per Page", "Schema - Org & Local", "OnPage - Headings",
                         "Page Inventory", "Action Plan", "Data Sources"]},
}


def in_scope(issue, scope):
    sc = SCOPES.get(scope)
    if not sc:
        return True
    if sc.get("only_categories") and issue["category"] not in sc["only_categories"] and not issue["issue"].startswith(tuple(sc.get("categories", ()))):
        return issue["area"] == "Technical" and issue["category"] == "Crawl"
    if issue["category"] in sc.get("exclude_categories", set()):
        return False
    return issue["area"] in sc["areas"] or issue["issue"].startswith(tuple(sc.get("categories", ()))) or \
        (issue["area"] == "Technical" and issue["category"] == "Crawl")


def where_for(title, fam):
    for prefix, opts in GUIDE:
        if title.startswith(prefix):
            return opts.get(fam) or opts.get("default") or next(iter(opts.values()))
    return ""


def code_snippet(kind, **k):
    """Copy-paste code for common fixes."""
    if kind == "redirect":
        src, dst, fam = k["src"], k["dst"], k.get("fam")
        path = urlparse(src).path
        if fam == "nextjs":
            return f"// next.config.js > redirects()\n{{ source: '{path}', destination: '{urlparse(dst).path}', permanent: true }}"
        if fam == "wordpress":
            return f"Rank Math > Redirections > Add New\nSource: {path}\nDestination: {dst}\nType: 301"
        return (f"# Apache .htaccess\nRedirect 301 {path} {dst}\n\n# Nginx\nlocation = {path} {{ return 301 {dst}; }}")
    if kind == "soft404":
        fam = k.get("fam")
        if fam == "nextjs":
            return "import { notFound } from 'next/navigation'\n// inside the page component / generateMetadata:\nif (!data) notFound()"
        if fam == "react":
            return ("# Netlify _redirects / Vercel: serve a real 404 for unknown paths, e.g. pre-render known routes and\n"
                    "# add a catch-all rule that returns /404.html with status 404")
        return "# Nginx\nerror_page 404 /404.html;\n# Ensure the app returns status 404 (not 200) for unmatched routes"
    if kind == "https":
        return ("# Apache .htaccess\nRewriteEngine On\nRewriteCond %{HTTPS} off\nRewriteRule ^(.*)$ https://%{HTTP_HOST}/$1 [L,R=301]\n\n"
                "# Nginx\nserver { listen 80; server_name _; return 301 https://$host$request_uri; }")
    if kind == "heading":
        return f"<!-- before -->\n<h{k['level']} class=\"...\">{k['text']}</h{k['level']}>\n<!-- after -->\n<p class=\"...\">{k['text']}</p>"
    if kind == "canonical":
        if k.get("fam") == "nextjs":
            return f"export const metadata = {{ alternates: {{ canonical: '{urlparse(k['url']).path}' }} }}"
        return f'<link rel="canonical" href="{k["url"]}">'
    return ""


def jsonld(obj):
    return '<script type="application/ld+json">\n' + json.dumps(obj, indent=1, ensure_ascii=False) + "\n</script>"


def is_broken(s):
    """Real errors only: HTTP 4xx/5xx or SSL/redirect-loop problems. No-response (server overload) is NOT 'broken'."""
    if s.get("unreachable"):
        return False
    err = s.get("error") or ""
    return (s.get("final_status") or 0) >= 400 or (bool(err) and "robots" not in err and "private" not in err)


def short(u, n=0):
    return u if not n or len(u or "") <= n else u[: n - 1] + "…"


def types_of(node):
    t = node.get("@type")
    return [str(x) for x in (t if isinstance(t, list) else [t]) if x]


def page_types(p):
    s = set()
    for n in p.get("schema", []):
        s.update(types_of(n))
    return s | set(p.get("microdata") or [])


def digits(s):
    return re.sub(r"\D", "", s or "")[-10:]


# ============================================================================ analysis
class Audit:
    def __init__(self, crawl, analysis):
        self.c = crawl
        self.a = analysis or {}
        self.pages = crawl["pages"]
        self.status = {s["url"]: s for s in crawl["status"]}
        self.sitemap = {s["url"]: s for s in crawl.get("sitemap_urls", [])}
        self.inlinks = crawl.get("inlinks", {})
        self.inlink_counts = crawl.get("inlink_counts", {})
        self.found_in = crawl.get("found_in", {})
        self.issues = []
        self.page_issues = defaultdict(list)
        self.fam = (crawl["meta"].get("platform") or {}).get("family", "custom")
        self.platform = (crawl["meta"].get("platform") or {}).get("label", "Unknown")
        self.scope = None

    def add(self, area, category, issue, severity, urls, why, fix, owner, evidence=""):
        urls = list(dict.fromkeys(u for u in urls if u))
        self.issues.append({"area": area, "category": category, "issue": issue, "severity": severity,
                            "count": len(urls) if urls else (1 if evidence else 0), "urls": urls,
                            "evidence": evidence, "why": why, "fix": fix, "owner": owner, "source": "Crawler"})
        for u in urls:
            self.page_issues[u].append(issue)

    # ---- technical
    def technical(self):
        sc = self.c.get("site_checks", {})
        start = self.c["meta"]["start_url"]
        if not self.pages:
            st = self.status.get(start, {})
            self.add("Technical", "Crawl", "Start URL could not be crawled - report is incomplete", "Critical", [start],
                     "Nothing else in this workbook is reliable until the page can be fetched.",
                     "Check the site is online. If it blocks bots/cloud servers (403/429/444), whitelist the audit "
                     "or run Screaming Frog from the office and use its export.", "SEO",
                     evidence=f"Status: {st.get('final_status') or st.get('status')} | Error: {st.get('error')}")
        redirects = [s for s in self.status.values() if s.get("chain")]
        chains = [s for s in redirects if len(s["chain"]) > 1]
        errors = [s for s in self.status.values() if is_broken(s)]
        unreach = [s["url"] for s in self.status.values() if s.get("unreachable")]
        if unreach:
            self.add("Technical", "Crawl", "URLs that gave no response during the crawl (not confirmed broken)", "Info", unreach,
                     "The server dropped/throttled connections; these may be fine. Not counted as broken.",
                     "Open a few in a browser. If they load, ignore. If many fail, check server rate-limiting / firewall.", "SEO")
        if redirects:
            self.add("Technical", "Redirects", "Internal URLs that redirect (3xx)", "Medium",
                     [s["url"] for s in redirects],
                     "Every redirect adds latency and wastes crawl budget; internal links should point to the final URL.",
                     "Update internal links/menus to the final URL shown in 'Tech - Redirects'. Keep the 301 for external visitors.",
                     "Developer")
        if chains:
            self.add("Technical", "Redirects", "Redirect chains (2+ hops)", "High", [s["url"] for s in chains],
                     "Chains lose link equity at each hop and slow crawling; Google stops after ~10 hops.",
                     "Point the first URL directly at the final destination (single 301).", "Developer")
        sm_redirect = [u for u in self.sitemap if self.status.get(u, {}).get("chain")]
        if sm_redirect:
            self.add("Technical", "Sitemap", "Redirecting URLs listed in the XML sitemap", "High", sm_redirect,
                     "Sitemaps should list only final, indexable 200 URLs; redirects send mixed signals.",
                     "Remove these from the sitemap (or replace with the final URL). In Yoast/Rank Math: noindex or delete the old page.",
                     "Developer")
        sm_err = [u for u in self.sitemap if (self.status.get(u, {}).get("final_status") or 0) >= 400]
        if sm_err:
            self.add("Technical", "Sitemap", "Broken (4xx/5xx) URLs listed in the XML sitemap", "High", sm_err,
                     "Google reports these as errors and loses trust in the sitemap.",
                     "Remove from the sitemap; restore the page or 301 it to the closest relevant page.", "Developer")
        if errors:
            linked = [s["url"] for s in errors if self.inlinks.get(s["url"])]
            self.add("Technical", "URL errors", "Broken internal URLs (4xx / 5xx / connection errors)",
                     "High" if linked else "Medium", [s["url"] for s in errors],
                     "Users and crawlers hit dead ends; link equity is wasted.",
                     "Fix or remove the links on the source pages listed in 'Tech - URL Errors', or 301 the URL to a relevant live page.",
                     "Developer")
        to_redirect = sorted({src for s in redirects for src in self.inlinks.get(s["url"], [])})
        if to_redirect:
            self.add("Technical", "Internal links", "Pages linking to redirecting URLs", "Low", to_redirect,
                     "Each click/crawl goes through an unnecessary redirect.",
                     "Edit the links on these pages to the final URL.", "Content team")
        dup_sm = [s["url"] for s in self.c.get("sitemap_urls", []) if s.get("listed", 1) > 1]
        if dup_sm:
            self.add("Technical", "Sitemap", "Duplicate URLs across sitemaps", "Low", dup_sm,
                     "Duplicates make sitemap counts unreliable.", "List each URL once.", "Developer")
        http = sc.get("http_root", {})
        if http and http.get("final_url") and start.startswith("https://") and \
                not str(http.get("final_url", "")).startswith("https://"):
            self.add("Technical", "HTTPS", "HTTP version does not redirect to HTTPS", "High", [http.get("url")],
                     "Duplicate http/https versions split ranking signals and are insecure.",
                     "Force a site-wide 301 from http:// to https:// (server or Cloudflare 'Always use HTTPS').", "Developer")
        alt = sc.get("alt_www_root", {})
        if alt and alt.get("final_status") == 200 and alt.get("hops", 0) == 0:
            self.add("Technical", "Canonical host", "Both www and non-www versions load (no redirect)", "High",
                     [alt.get("url")], "Two copies of the site compete in search.",
                     "301 the non-preferred host to the preferred one.", "Developer")
        if sc.get("soft_404_test", {}).get("status") == 200:
            self.add("Technical", "Status codes", "Non-existent URLs return 200 (soft 404)", "High",
                     [], "Google may index junk URLs; real 404s are hidden.",
                     "Make the server return a real 404 status for missing pages.", "Developer",
                     evidence="A random URL returned HTTP 200.")
        if not sc.get("robots_txt", {}).get("found") and self.c["meta"]["mode"] != "single page":
            self.add("Technical", "Robots", "robots.txt missing or unreachable", "Medium", [],
                     "Crawlers get no guidance and no sitemap location.",
                     "Add /robots.txt with a 'Sitemap:' line.", "Developer", evidence="robots.txt not found")
        if not self.c.get("sitemaps") and self.c["meta"]["mode"] != "single page":
            self.add("Technical", "Sitemap", "No XML sitemap found", "High", [],
                     "Google discovers pages more slowly.", "Enable the sitemap in Yoast / Rank Math and submit it in Search Console.",
                     "Developer", evidence="Checked robots.txt, /sitemap.xml, /sitemap_index.xml, /wp-sitemap.xml")
        # indexability
        noindex = [p["url"] for p in self.pages if "noindex" in ((p.get("meta_robots") or "") + (p.get("x_robots_tag") or "")).lower()]
        noindex_in_sm = [u for u in noindex if u in self.sitemap]
        if noindex_in_sm:
            self.add("Technical", "Indexability", "Noindex pages listed in the sitemap", "High", noindex_in_sm,
                     "Contradictory signal: 'index this' (sitemap) vs 'do not index' (robots).",
                     "Remove from the sitemap, or remove noindex if the page should rank.", "Developer")
        canon_bad, canon_other, canon_missing = [], [], []
        for p in self.pages:
            c = p.get("canonical")
            if not c:
                canon_missing.append(p["url"])
            elif c.rstrip("/") != p["url"].rstrip("/"):
                canon_other.append(p["url"])
                st = self.status.get(c, {})
                if st and ((st.get("final_status") or 0) >= 300 or st.get("chain")):
                    canon_bad.append(p["url"])
        if canon_missing:
            self.add("Technical", "Canonical", "Missing canonical tag", "Medium", canon_missing,
                     "Without a canonical, parameter/duplicate URLs can be indexed instead.",
                     "Enable self-referencing canonicals (Yoast/Rank Math do this by default).", "Developer")
        if canon_other:
            self.add("Technical", "Canonical", "Canonical points to a different URL", "Info", canon_other,
                     "Fine if intentional (duplicates); a mistake if the page should rank itself.",
                     "Check each one in 'Tech - Indexability'.", "SEO")
        if canon_bad:
            self.add("Technical", "Canonical", "Canonical points to a redirecting or broken URL", "High", canon_bad,
                     "Google ignores invalid canonicals.", "Set the canonical to the final 200 URL.", "Developer")
        slow = [p["url"] for p in self.pages if (p.get("elapsed_ms") or 0) > 1500]
        if slow:
            self.add("Technical", "Performance", "Slow server response (>1.5 s HTML download)", "Medium", slow,
                     "Slow TTFB hurts LCP and crawl rate.",
                     "Enable full-page caching (LiteSpeed/WP Rocket/Cloudflare APO) and check hosting.", "Developer")
        mixed = [p["url"] for p in self.pages if p.get("mixed_content")]
        if mixed:
            self.add("Technical", "HTTPS", "Mixed content (http:// resources on https pages)", "High", mixed,
                     "Browsers block or warn; the padlock breaks.", "Change resource URLs to https:// (Better Search Replace).", "Developer")
        no_vp = [p["url"] for p in self.pages if not p.get("viewport")]
        if no_vp:
            self.add("Technical", "Mobile", "Missing viewport meta tag", "High", no_vp,
                     "Page is not mobile-friendly; Google indexes mobile-first.",
                     "Add <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">.", "Developer")
        no_lang = [p["url"] for p in self.pages if not p.get("lang")]
        if no_lang:
            self.add("Technical", "HTML", "Missing <html lang> attribute", "Low", no_lang,
                     "Helps search engines and screen readers detect the language.", "Set lang=\"en\" (or correct locale) in the theme.", "Developer")
        # trailing slash inconsistency
        variants = defaultdict(set)
        for u in list(self.inlinks):
            pth = urlparse(u).path
            if pth not in ("", "/"):
                variants[u.rstrip("/")].add(u)
        both = [k for k, v in variants.items() if len(v) > 1]
        if both:
            self.add("Technical", "URL structure", "Same URL linked with and without trailing slash", "Low", both,
                     "Creates duplicate URLs / extra redirects.", "Pick one format and keep internal links consistent.", "Developer")
        upper = [u for u in self.inlinks if re.search(r"[A-Z]", urlparse(u).path)]
        if upper:
            self.add("Technical", "URL structure", "Uppercase letters in internal URLs", "Low", upper,
                     "URLs are case-sensitive; mixed case creates duplicates.", "Use lowercase URLs; 301 uppercase versions.", "Developer")
        # orphans (only meaningful within the crawl)
        crawled = {p["url"] for p in self.pages}
        orphans = [u for u in self.sitemap if u in crawled and not self.inlinks.get(u) and
                   u.rstrip("/") != self.c["meta"]["start_url"].rstrip("/")]
        if orphans:
            self.add("Technical", "Internal links", "Orphan pages (in sitemap, no internal links found in crawl)", "Medium", orphans,
                     "Pages without internal links get little authority and are crawled less.",
                     "Link to them from relevant pages/menus, or remove them if obsolete.", "Content team")

    # ---- on-page
    def onpage(self):
        P = self.pages
        utility = ("Utility page (thank-you / cart / account)", "Legal / policy page")
        missing_t = [p["url"] for p in P if not p.get("title")]
        long_t = [p["url"] for p in P if p.get("title") and len(p["title"]) > 60]
        short_t = [p["url"] for p in P if p.get("title") and len(p["title"]) < 30 and p["page_type"] not in utility]
        if missing_t:
            self.add("On-Page", "Titles", "Missing or empty <title>", "High", missing_t,
                     "The title is the strongest on-page ranking signal and the SERP headline.",
                     "Write a unique 50-60 character title with the main keyword first.", "Content team")
        if long_t:
            self.add("On-Page", "Titles", "Title longer than 60 characters (truncated in Google)", "Medium", long_t,
                     "Google cuts titles at ~580px (~60 chars); the end (often the brand/CTA) is lost.",
                     "Shorten to 50-60 characters: 'Primary Keyword - Benefit | Brand'.", "Content team")
        if short_t:
            self.add("On-Page", "Titles", "Title shorter than 30 characters", "Low", short_t,
                     "Under-uses the most valuable on-page space.", "Expand with keyword + benefit/location.", "Content team")
        dups = defaultdict(list)
        for p in P:
            if p.get("title"):
                dups[p["title"].strip().lower()].append(p["url"])
        dup_t = [u for v in dups.values() if len(v) > 1 for u in v]
        if dup_t:
            self.add("On-Page", "Titles", "Duplicate titles", "Medium", dup_t,
                     "Pages compete with each other (cannibalisation).", "Make each title unique.", "Content team")
        missing_m = [p["url"] for p in P if not p.get("meta_description") and p["page_type"] not in utility]
        long_m = [p["url"] for p in P if p.get("meta_description") and len(p["meta_description"]) > 160]
        short_m = [p["url"] for p in P if p.get("meta_description") and len(p["meta_description"]) < 70]
        if missing_m:
            self.add("On-Page", "Meta description", "Missing meta description", "Medium", missing_m,
                     "Google writes its own snippet, often less persuasive -> lower CTR.",
                     "Write a 140-155 character description with keyword + CTA.", "Content team")
        if long_m:
            self.add("On-Page", "Meta description", "Meta description over 160 characters", "Low", long_m,
                     "Truncated in results.", "Trim to 140-155 characters.", "Content team")
        if short_m:
            self.add("On-Page", "Meta description", "Meta description under 70 characters", "Low", short_m,
                     "Too short to persuade.", "Expand to 140-155 characters.", "Content team")
        dm = defaultdict(list)
        for p in P:
            if p.get("meta_description"):
                dm[p["meta_description"].strip().lower()].append(p["url"])
        dup_m = [u for v in dm.values() if len(v) > 1 for u in v]
        if dup_m:
            self.add("On-Page", "Meta description", "Duplicate meta descriptions", "Low", dup_m,
                     "Looks templated to Google; lower CTR.", "Write unique descriptions.", "Content team")
        # headings
        no_h1, multi_h1, skipped, first_not_h1, empty_h, h1_in_tpl = [], [], [], [], [], []
        for p in P:
            hs = p.get("headings", [])
            content = [h for h in hs if h["region"] in ("main", "body")]
            h1 = [h for h in hs if h["level"] == 1]
            if not h1:
                no_h1.append(p["url"])
            if len(h1) > 1:
                multi_h1.append(p["url"])
            if any(h["level"] == 1 and h["region"] in ("header", "footer", "nav", "aside") for h in hs):
                h1_in_tpl.append(p["url"])
            if content and content[0]["level"] != 1 and h1:
                first_not_h1.append(p["url"])
            prev = 0
            for h in content:
                if prev and h["level"] > prev + 1:
                    skipped.append(p["url"])
                    break
                prev = h["level"]
            if any(h["empty"] for h in hs):
                empty_h.append(p["url"])
        if no_h1:
            self.add("On-Page", "Headings", "Missing H1", "High", no_h1,
                     "The H1 tells users and Google the page's main topic.",
                     "Add exactly one H1 containing the primary keyword.", "Content team")
        if multi_h1:
            self.add("On-Page", "Headings", "Multiple H1 tags", "Medium", multi_h1,
                     "Dilutes the main topic signal; usually a theme/builder mistake.",
                     "Keep one H1; change the others to H2 (Elementor: Heading widget > HTML Tag).", "Content team")
        if h1_in_tpl:
            self.add("On-Page", "Headings", "H1 inside header/footer/menu (site-wide template)", "High", h1_in_tpl,
                     "Every page gets the same H1 (e.g. the logo), competing with the real one.",
                     "Change the template element to a <div>/<p> (logo should not be an H1 on inner pages).", "Developer")
        if skipped:
            self.add("On-Page", "Headings", "Heading hierarchy skips levels (e.g. H2 -> H4)", "Low", skipped,
                     "Broken outline hurts accessibility and topic structure.",
                     "Use H1 > H2 > H3 in order; see the outline in 'OnPage - Headings'.", "Content team")
        if first_not_h1:
            self.add("On-Page", "Headings", "Content starts with H2-H6 before the H1", "Low", first_not_h1,
                     "Main topic heading should come first in the content.", "Move the H1 to the top of the content.", "Content team")
        if empty_h:
            self.add("On-Page", "Headings", "Empty heading tags", "Low", empty_h,
                     "Empty headings confuse screen readers and the outline.", "Remove them or add text.", "Developer")
        tpl = [p["url"] for p in P if any(h["region"] in ("header", "footer", "nav", "aside") for h in p.get("headings", []))]
        if tpl:
            self.add("On-Page", "Headings", "Headings used in header / footer / menu / sidebar", "Medium", tpl,
                     "Template headings repeat on every page and pollute each page's outline "
                     "(e.g. 'Quick Links' as H4 on 500 pages).",
                     "Change these to <p>/<div>/<span> styled with CSS (Elementor: Heading widget > HTML Tag = div/p). "
                     "See 'OnPage - Header-Footer H' for the exact elements.", "Developer")
        thin = [p["url"] for p in P if p["word_count"] < 300 and p["page_type"] in
                ("Landing / service page", "Blog post / article", "Product page", "Location page", "Homepage")]
        if thin:
            self.add("On-Page", "Content", "Thin content (<300 words in main content)", "Medium", thin,
                     "Too little content to rank for competitive queries or show expertise.",
                     "Expand with FAQs, benefits, process, proof (reviews/case studies).", "Content team")
        no_og = [p["url"] for p in P if not p.get("og_title") or not p.get("og_image")]
        if no_og:
            self.add("On-Page", "Social", "Missing Open Graph title/image", "Low", no_og,
                     "Shares on WhatsApp/Facebook/LinkedIn show no image or a wrong title.",
                     "Set OG title + 1200x630 image (Yoast/Rank Math Social tab).", "Content team")
        # images
        imgs = self.c.get("images", [])
        big = [i for i in imgs if (i.get("bytes") or 0) > 100 * 1024]
        huge = [i for i in imgs if (i.get("bytes") or 0) > 500 * 1024]
        if big:
            self.add("On-Page", "Images", "Images larger than 100 KB", "High" if huge else "Medium",
                     [i["url"] for i in big],
                     "Heavy images are the #1 cause of slow LCP on WordPress sites.",
                     "Compress and convert to WebP/AVIF (ShortPixel/Imagify/EWWW), resize to display size.", "Developer")
        legacy = [i["url"] for i in big if urlparse(i["url"]).path.lower().endswith(LEGACY_IMG)]
        if legacy:
            self.add("On-Page", "Images", "Large images in legacy format (JPG/PNG) - not WebP/AVIF", "Medium", legacy,
                     "WebP is typically 25-35% smaller at the same quality.", "Serve WebP/AVIF versions.", "Developer")
        no_alt_pages, no_alt_imgs = set(), set()
        for p in P:
            for im in p.get("images", []):
                if im["region"] in ("main", "body") and (not im["has_alt_attr"] or not (im["alt"] or "").strip()):
                    no_alt_pages.add(p["url"])
                    no_alt_imgs.add(im["src"])
        if no_alt_imgs:
            self.add("On-Page", "Images", "Content images missing alt text", "Medium", sorted(no_alt_pages),
                     "Alt text is needed for accessibility and image search.",
                     "Add descriptive alt text (what the image shows + context), not file names.", "Content team",
                     evidence=f"{len(no_alt_imgs)} unique images")
        broken_imgs = [i["url"] for i in imgs if (i.get("status") or 0) >= 400]
        if broken_imgs:
            self.add("On-Page", "Images", "Broken images (4xx/5xx)", "High", broken_imgs,
                     "Broken images look unprofessional and waste requests.", "Re-upload or remove them.", "Developer")
        nodim = sorted({p["url"] for p in P for im in p.get("images", [])
                        if im["region"] in ("main", "body") and not (im.get("width") and im.get("height"))})
        if nodim:
            self.add("On-Page", "Images", "Images without width/height attributes", "Low", nodim,
                     "Causes layout shift (CLS).", "Set width and height on <img> (most builders do with 'image size').", "Developer")

    # ---- schema
    def schema(self):
        P = self.pages
        invalid = [p["url"] for p in P if p.get("schema_errors")]
        if invalid:
            self.add("Schema", "Validity", "Invalid JSON-LD (cannot be parsed)", "High", invalid,
                     "Google ignores broken structured data entirely.", "Fix the JSON syntax; test in validator.schema.org.", "Developer")
        none = [p["url"] for p in P if not page_types(p)]
        if none:
            self.add("Schema", "Coverage", "Pages with no structured data", "Medium", none,
                     "No eligibility for rich results; weaker entity understanding.", "Add schema via Yoast/Rank Math or custom JSON-LD.", "Developer")
        missing_exp = []
        for p in P:
            have = {t.lower() for t in page_types(p)}
            for exp in EXPECTED_SCHEMA.get(p["page_type"], []):
                opts = {o.lower() for o in exp.split("|")}
                if exp == "LocalBusiness":
                    opts |= LOCAL_TYPES
                if not (have & opts):
                    missing_exp.append(p["url"])
                    break
        if missing_exp:
            self.add("Schema", "Coverage", "Schema type expected for this page type is missing", "Medium", missing_exp,
                     "Each page type has a schema that unlocks rich results (Product, Article, LocalBusiness, Breadcrumb).",
                     "See 'Schema - Per Page' column 'Missing' for exactly which type to add.", "Developer")
        org_pages = [p for p in P if {t.lower() for t in page_types(p)} & ORG_TYPES]
        home = next((p for p in P if p["page_type"] == "Homepage"), None)
        if not org_pages:
            self.add("Schema", "Organization", "No Organization schema found", "High",
                     [home["url"]] if home else [],
                     "Organization schema defines your brand entity (logo, social profiles) for Google/AI search.",
                     "Add Organization with name, url, logo, sameAs (social profiles), contactPoint.", "Developer")
        else:
            names = set()
            for p in org_pages:
                for n in p["schema"]:
                    if {t.lower() for t in types_of(n)} & (ORG_TYPES - LOCAL_TYPES):
                        names.add(str(n.get("name", "")).strip())
            names.discard("")
            if len(names) > 1:
                self.add("Schema", "Organization", "Organization name differs across pages / multiple entities",
                         "Medium", [p["url"] for p in org_pages],
                         "Conflicting entity data weakens the Knowledge Graph signal.",
                         "Output one Organization node with a fixed @id site-wide.", "Developer",
                         evidence="Names found: " + "; ".join(sorted(n for n in names if n)))
        local = [p for p in P if {t.lower() for t in page_types(p)} & LOCAL_TYPES]
        if local:
            phones = set()
            for p in local:
                for n in p["schema"]:
                    if {t.lower() for t in types_of(n)} & LOCAL_TYPES and n.get("telephone"):
                        phones.add(digits(str(n["telephone"])))
        rated = [p["url"] for p in P for n in p.get("schema", [])
                 if ({t.lower() for t in types_of(n)} & ORG_TYPES) and (n.get("aggregateRating") or n.get("review"))]
        if rated:
            self.add("Schema", "Reviews", "Review/rating markup on the business's own Organization/LocalBusiness", "Medium",
                     sorted(set(rated)),
                     "Google does not show review stars for self-serving reviews on an Organization/LocalBusiness; "
                     "if the ratings are not real, visible reviews on the page, it breaks Google's structured-data guidelines.",
                     "Remove aggregateRating/review from the Organization/LocalBusiness node, or move genuine ratings to "
                     "Course/Product/Service items that show the reviews on the page.", "Developer")

    # ---- off-page signals visible from the site
    def offpage(self):
        ext = self.c.get("external_links", [])
        broken = [e["url"] for e in ext if (e.get("status") or 0) >= 400 or (e.get("error") and e.get("status") is None)]
        if broken:
            self.add("Off-Page", "Outbound links", "Broken external links", "Medium", broken,
                     "Dead outbound links signal an unmaintained site.", "Update or remove them (sources in 'OffPage - Links & Profiles').", "Content team")
        social = sorted({e["url"] for e in ext if any(s in urlparse(e["url"]).netloc for s in SOCIAL)})
        same_as = set()
        for p in self.pages:
            for n in p.get("schema", []):
                sa = n.get("sameAs")
                for s in (sa if isinstance(sa, list) else [sa] if sa else []):
                    same_as.add(str(s).rstrip("/").lower())
        not_in_sameas = [s for s in social if s.rstrip("/").lower() not in same_as
                         and urlparse(s).path.strip("/") and "share" not in s and "intent" not in s]
        if not_in_sameas:
            self.add("Off-Page", "Entity / brand", "Social profiles linked on site but missing from Organization sameAs", "Low",
                     not_in_sameas, "sameAs connects your brand entity to its profiles (Knowledge Panel, AI answers).",
                     "Add these profile URLs to Organization 'sameAs'.", "Developer")
        tels = Counter(digits(t) for p in self.pages for t in p.get("tel_links", []) if digits(t))
        schema_tels = {digits(str(n.get("telephone"))) for p in self.pages for n in p.get("schema", []) if n.get("telephone")}
        if tels and schema_tels and not (set(tels) & schema_tels):
            self.add("Off-Page", "NAP", "Phone number on site differs from schema telephone", "Medium", [],
                     "NAP (Name/Address/Phone) must match exactly across site, schema and Google Business Profile.",
                     "Use one phone format everywhere and match the GBP listing.", "SEO",
                     evidence=f"tel: links {sorted(tels)} vs schema {sorted(schema_tels)}")

    def cwv(self):
        psi = self.c.get("pagespeed") or []
        ok = [x for x in psi if not x.get("error")]
        if not ok:
            self.add("Technical", "Core Web Vitals", "Core Web Vitals not measured", "Info", [],
                     "Speed is a ranking factor and affects conversions.",
                     "Add a free Google API key as PAGESPEED_API_KEY in the cloud environment, then re-run.", "Developer",
                     evidence="; ".join(x.get("error", "") for x in psi)[:200] or "No PageSpeed key configured")
            return
        poor_lab = [x["url"] for x in ok if x.get("perf_score", 100) < 50]
        if poor_lab:
            self.add("Technical", "Core Web Vitals", "Poor mobile performance score (<50)", "High", poor_lab,
                     "Slow mobile pages lose rankings and visitors.",
                     "Fix the top opportunities in 'Tech - Core Web Vitals' (images, render-blocking CSS/JS, caching).", "Developer")
        for metric, label, limit in (("lcp", "LCP", 2500), ("inp", "INP", 200)):
            bad = [x["url"] for x in ok if (x.get("field_url") or {}).get(f"{metric}_cat") in ("SLOW", "AVERAGE")]
            if bad:
                self.add("Technical", "Core Web Vitals", f"Real-user {label} not 'Good' (target < {limit} ms)", "High", bad,
                         "Google uses real-user Core Web Vitals (CrUX) as a ranking signal.",
                         "See 'Tech - Core Web Vitals' for the measured value and top fixes.", "Developer")
        bad_cls = [x["url"] for x in ok if (x.get("field_url") or {}).get("cls_cat") in ("SLOW", "AVERAGE")]
        if bad_cls:
            self.add("Technical", "Core Web Vitals", "Real-user CLS not 'Good' (target < 0.1)", "Medium", bad_cls,
                     "Layout shifts frustrate users.", "Set image/embed dimensions; reserve space for banners/fonts.", "Developer")

    # ---- per-URL fix list (what to change, where, current -> new)
    def brand(self):
        for p in self.pages:
            for n in p.get("schema", []):
                if {t.lower() for t in types_of(n)} & ORG_TYPES and n.get("name"):
                    return str(n["name"]).strip()
        return self.c["meta"]["domain"].split(".")[0].title()

    def fix_rows(self):
        fam = self.fam
        rows, add = [], None
        crawled = {p["url"]: p for p in self.pages}
        live_paths = {urlparse(u).path: u for u in crawled}
        brand = self.brand()
        cap = 4000

        def add(issue, sev, page, element, current, change, where, code="", owner=""):
            if len(rows) < cap:
                rows.append([issue, sev, page, element, current, change, where, code, owner])

        def anchor_on(src, href):
            for ln in crawled.get(src, {}).get("links", []):
                if ln["href"] == href:
                    return ln["text"] or "(image/icon link)", ln["region"]
            return "", ""

        def link_rows(issue_label, sev_val, target, change, where_key):
            """One row per source page; links in header/footer/nav on many pages collapse to one template row."""
            srcs = self.inlinks.get(target, [])
            tpl = defaultdict(list)
            body = []
            for src in srcs:
                text, region = anchor_on(src, target)
                (tpl[(region, text)] if region in ("header", "footer", "nav", "aside") else body).append((src, text, region))
            n_pages = max(len(self.pages), 1)
            for (region, text), lst in tpl.items():
                if len(lst) >= 3 or len(lst) >= 0.5 * n_pages:
                    add(issue_label, sev_val, f"All pages with the {region} ({self.inlink_counts.get(target, len(lst))} pages)",
                        f"{region.title()} template: <a> '{text}'", target, change, where_for(where_key, fam), owner="Developer")
                else:
                    body.extend(lst)
            for src, text, region in body[:20]:
                add(issue_label, sev_val, src, f"<a> link '{text}' in {region or 'content'}", target, change,
                    where_for(where_key, fam), owner="Content team" if region in ("main", "body", "") else "Developer")

        def sev(prefix):
            return next((i["severity"] for i in self.issues if i["issue"].startswith(prefix)), "Medium")

        # broken internal links -> each source page
        for s in self.status.values():
            fs = s.get("final_status") or 0
            if not is_broken(s):
                continue
            match = difflib.get_close_matches(urlparse(s["url"]).path, list(live_paths), n=1, cutoff=0.75)
            suggestion = f"closest live page (verify): {live_paths[match[0]]}" if match else "remove the link, or point it to the right live page"
            link_rows("Broken internal link", sev("Broken internal URLs"), s["url"],
                      f"Returns {fs or s.get('error')}. Change link to {suggestion}", "Broken internal URLs")
            if s["url"] in self.sitemap:
                add("Broken URL in sitemap", sev("Broken (4xx/5xx) URLs listed"), ", ".join(self.sitemap[s["url"]]["sitemaps"]),
                    "<loc> entry", s["url"], "Remove from sitemap" + (f" (or 301 to the {suggestion})" if match else ""),
                    where_for("Broken (4xx/5xx) URLs listed in the XML sitemap", fam), owner="Developer")
        # redirects
        for s in self.status.values():
            if not s.get("chain"):
                continue
            final = s.get("final_url") or ""
            if s["url"] in self.sitemap:
                add("Redirecting URL in sitemap", sev("Redirecting URLs listed"), ", ".join(self.sitemap[s["url"]]["sitemaps"]),
                    "<loc> entry", s["url"], f"Replace with: {final}", where_for("Redirecting URLs listed in the XML sitemap", fam),
                    owner="Developer")
            if len(s["chain"]) > 1:
                add("Redirect chain", sev("Redirect chains"), s["url"], "Redirect rule",
                    " -> ".join([h["url"] for h in s["chain"]] + [final]), f"One 301: {s['url']} -> {final}",
                    where_for("Redirect chains", fam), code_snippet("redirect", src=s["url"], dst=final, fam=fam), "Developer")
            link_rows("Internal link to redirect", sev("Internal URLs that redirect"), s["url"],
                      f"Change link to: {final}", "Internal URLs that redirect")
        # titles
        tcount = Counter((p.get("title") or "").strip().lower() for p in self.pages if p.get("title"))
        for p in self.pages:
            t = (p.get("title") or "").strip()
            h1 = next((h["text"] for h in p["headings"] if h["level"] == 1 and h["text"]), "")
            if not t:
                add("Missing title", sev("Missing or empty <title>"), p["url"], "<title>", "(empty)",
                    f"Suggested: {short((h1 or p['url']) + ' | ' + brand, 60)}", where_for("Missing or empty <title>", fam), owner="Content team")
            elif len(t) > 60 or tcount[t.lower()] > 1:
                core = re.split(r"\s[|\-–—:]\s", t)[0].strip()
                if tcount[t.lower()] > 1 and h1 and h1.lower() not in t.lower():
                    core = h1
                cand = f"{core} | {brand}" if len(core) + len(brand) + 3 <= 60 else core
                if len(cand) > 60:
                    cand = cand[:60].rsplit(" ", 1)[0]
                why = (f"{len(t)} chars" if len(t) > 60 else "") + (" + duplicate" if tcount[t.lower()] > 1 else "")
                add("Title too long / duplicate", sev("Title longer") if len(t) > 60 else sev("Duplicate titles"), p["url"],
                    "<title>", f"{t} ({why.strip(' +')})", f"Suggested ({len(cand)} chars): {cand}  [review before use]",
                    where_for("Title ", fam), owner="Content team")
        # meta descriptions
        mcount = Counter((p.get("meta_description") or "").strip().lower() for p in self.pages if p.get("meta_description"))
        for p in self.pages:
            d = (p.get("meta_description") or "").strip()
            h1 = next((h["text"] for h in p["headings"] if h["level"] == 1 and h["text"]), p.get("title") or "")
            if not d and p["page_type"] not in ("Utility page (thank-you / cart / account)", "Legal / policy page"):
                add("Missing meta description", sev("Missing meta description"), p["url"], '<meta name="description">', "(missing)",
                    f"Write 140-155 chars: what '{short(h1, 60)}' offers + benefit + call to action",
                    where_for("Missing meta description", fam), owner="Content team")
            elif d and (len(d) > 160 or mcount[d.lower()] > 1):
                cut = d if len(d) <= 155 else d[:155].rsplit(" ", 1)[0].rstrip(",;:-") + "."
                change = f"Suggested: {cut}" if mcount[d.lower()] == 1 else \
                    f"Rewrite uniquely around '{short(h1, 60)}' (currently shared by {mcount[d.lower()]} pages)"
                add("Meta description too long / duplicate", "Low", p["url"], '<meta name="description">',
                    f"{short(d, 200)} ({len(d)} chars)", change, where_for("Meta description", fam), owner="Content team")
        # headings
        for p in self.pages:
            hs = p["headings"]
            h1s = [h for h in hs if h["level"] == 1]
            if not h1s:
                first = next((h for h in hs if h["region"] in ("main", "body")), None)
                add("Missing H1", sev("Missing H1"), p["url"], f"First content heading: H{first['level']} '{first['text']}'" if first else "No headings",
                    "(no H1)", f"Make the main heading an H1: '{short(first['text'] if first else re.split(r' [|-] ', p.get('title') or '')[0], 70)}'",
                    where_for("Missing H1", fam), owner="Content team")
            elif len(h1s) > 1:
                add("Multiple H1", sev("Multiple H1"), p["url"], " | ".join(f"H1 '{short(h['text'], 40)}' ({h['region']})" for h in h1s),
                    f"{len(h1s)} H1 tags", f"Keep H1: '{short(h1s[0]['text'], 50)}'. Change the others to H2.",
                    where_for("Multiple H1", fam), owner="Content team")
            prev = 0
            for h in [x for x in hs if x["region"] in ("main", "body")]:
                if prev and h["level"] > prev + 1:
                    add("Heading level skipped", "Low", p["url"], f"H{h['level']} '{short(h['text'], 50)}'", f"H{prev} -> H{h['level']}",
                        f"Change to H{prev + 1}", where_for("Heading hierarchy skips", fam), owner="Content team")
                    break
                prev = h["level"]
        # header/footer headings (template level)
        groups = defaultdict(set)
        for p in self.pages:
            for h in p["headings"]:
                if h["region"] in ("header", "footer", "nav", "aside"):
                    groups[(h["region"], h["level"], h["text"] or "(empty)")].add(p["url"])
        for (region, lvl, text), urls in sorted(groups.items(), key=lambda kv: -len(kv[1])):
            add("Heading in header/footer", sev("Headings used in header"),
                "All pages (site-wide template)" if len(urls) > 1 and len(urls) >= 0.6 * len(self.pages) else "; ".join(sorted(urls)[:3]),
                f"{region.title()} template: H{lvl} '{text}'", f"<h{lvl}> ({len(urls)} pages)", "<p> or <div> with the same class",
                where_for("Headings used in header", fam), code_snippet("heading", level=lvl, text=text), "Developer")
        # images
        img_alt_pages = defaultdict(list)
        for p in self.pages:
            for im in p["images"]:
                if im["region"] in ("main", "body") and (not im["has_alt_attr"] or not (im.get("alt") or "").strip()):
                    img_alt_pages[im["src"]].append(p["url"])
        for i in sorted(self.c.get("images", []), key=lambda x: -(x.get("bytes") or 0)):
            kb = (i.get("bytes") or 0) / 1024
            if kb > 100:
                ext = os.path.splitext(urlparse(i["url"]).path)[1].lower()
                add("Image over 100 KB", sev("Images larger than 100 KB"), "; ".join(i.get("pages", [])[:3]) + (
                    f" (+{i['page_count'] - 3} more)" if i.get("page_count", 0) > 3 else ""), f"Image file: {i['url']}",
                    f"{kb:.0f} KB {ext.lstrip('.') or i.get('type', '')}",
                    "Compress to < 100 KB" + (" and convert to WebP" if ext in LEGACY_IMG else "") + "; resize to the displayed width",
                    where_for("Images larger than 100 KB", fam), owner="Developer")
        for src, pages in img_alt_pages.items():
            name = re.sub(r"[-_]+|\d{2,}x\d{2,}|\.\w+$", " ", os.path.basename(urlparse(src).path)).strip()
            add("Image missing alt text", sev("Content images missing alt"), "; ".join(pages[:3]) + (f" (+{len(pages) - 3} more)" if len(pages) > 3 else ""),
                f"<img src='{short(src, 90)}'>", "alt missing/empty",
                f"Describe the image, e.g. '{name or 'what the image shows'}' + context  [review]",
                where_for("Content images missing alt", fam), owner="Content team")
        # canonicals
        for p in self.pages:
            c = p.get("canonical")
            if not c:
                add("Missing canonical", sev("Missing canonical"), p["url"], '<link rel="canonical">', "(missing)", f"Self-canonical: {p['url']}",
                    where_for("Missing canonical", fam), code_snippet("canonical", url=p["url"], fam=fam), "Developer")
            else:
                st = self.status.get(c, {})
                if c.rstrip("/") != p["url"].rstrip("/") and st and ((st.get("final_status") or 0) >= 300 or st.get("chain")):
                    add("Canonical to redirect/broken URL", sev("Canonical points to a redirecting"), p["url"], '<link rel="canonical">', c,
                        f"Set to: {st.get('final_url') or p['url']}", where_for("Canonical points to a redirecting", fam),
                        code_snippet("canonical", url=st.get("final_url") or p["url"], fam=fam), "Developer")
            if "noindex" in ((p.get("meta_robots") or "") + (p.get("x_robots_tag") or "")).lower() and p["url"] in self.sitemap:
                add("Noindex page in sitemap", sev("Noindex pages listed"), p["url"], "meta robots + sitemap", "noindex AND in sitemap",
                    "Remove from sitemap, OR remove noindex if it should rank", where_for("Noindex pages listed in the sitemap", fam), owner="SEO")
        # orphans -> suggest hub
        crawled_set = set(crawled)
        for u in self.sitemap:
            if u in crawled_set and not self.inlinks.get(u) and u.rstrip("/") != self.c["meta"]["start_url"].rstrip("/"):
                parts = urlparse(u).path.rstrip("/").split("/")
                hub = None
                while len(parts) > 1 and not hub:
                    parts = parts[:-1]
                    hub = live_paths.get("/".join(parts) + "/") or live_paths.get("/".join(parts) or "/")
                add("Orphan page", sev("Orphan pages"), u, "Internal links", "0 internal links found",
                    f"Add a contextual link from: {hub or 'a relevant category/hub page'}", where_for("Orphan pages", fam), owner="Content team")
        # site-level
        sc = self.c.get("site_checks", {})
        if sc.get("soft_404_test", {}).get("status") == 200:
            add("Soft 404", sev("Non-existent URLs return 200"), "Whole site (routing / server)", "HTTP status for unknown URLs",
                "200 OK", "404 Not Found", where_for("Non-existent URLs return 200", fam), code_snippet("soft404", fam=fam), "Developer")
        h = sc.get("http_root", {})
        if self.c["meta"]["start_url"].startswith("https://") and h and h.get("final_url") and not str(h["final_url"]).startswith("https://"):
            add("HTTP not redirected", sev("HTTP version"), h["url"], "Server/CDN", f"-> {h.get('final_url')}", "301 to https://",
                where_for("HTTP version does not redirect", fam), code_snippet("https"), "Developer")
        # schema snippets
        org = None
        for p in self.pages:
            for n in p.get("schema", []):
                if {t.lower() for t in types_of(n)} & ORG_TYPES:
                    org = n
                    break
            if org:
                break
        socials = sorted({e["url"] for e in self.c.get("external_links", []) if any(s in urlparse(e["url"]).netloc for s in SOCIAL)
                          and urlparse(e["url"]).path.strip("/") and "share" not in e["url"] and "intent" not in e["url"]})
        home = next((p for p in self.pages if p["page_type"] == "Homepage"), self.pages[0] if self.pages else None)
        dom = self.c["meta"]["start_url"].split("/")[0] + "//" + urlparse(self.c["meta"]["start_url"]).netloc
        tel = next((t for p in self.pages for t in p.get("tel_links", [])), "")
        org_snip = {"@context": "https://schema.org", "@type": (types_of(org)[0] if org else "Organization"),
                    "@id": f"{dom}/#organization", "name": (org or {}).get("name") or brand, "url": dom + "/",
                    "logo": (org or {}).get("logo") or (home or {}).get("og_image") or f"{dom}/logo.png",
                    "sameAs": socials or ["https://www.facebook.com/<page>", "https://www.instagram.com/<handle>"]}
        if tel:
            org_snip["contactPoint"] = {"@type": "ContactPoint", "telephone": tel, "contactType": "customer service"}
        org_issue = next((i for i in self.issues if i["issue"].startswith(("No Organization schema", "Organization name differs",
                                                                              "Social profiles linked"))), None)
        if org_issue:
            add("Organization schema", org_issue["severity"], "All pages (site-wide, once)", "<head> JSON-LD",
                org_issue["issue"], "Use this single Organization node (fill placeholders)",
                where_for("No Organization schema", fam), jsonld(org_snip), "Developer")
        rated = [i for i in self.issues if i["issue"].startswith("Review/rating markup")]
        if rated:
            add("Self-serving rating markup", rated[0]["severity"], f"{rated[0]['count']} pages (template)", "JSON-LD aggregateRating / review",
                "aggregateRating on own Organization/LocalBusiness", "Delete the aggregateRating and review properties",
                where_for("Review/rating markup", fam), owner="Developer")
        n_schema = 0
        for p in self.pages:
            have = {t.lower() for t in page_types(p)}
            for exp in EXPECTED_SCHEMA.get(p["page_type"], []):
                opts = {o.lower() for o in exp.split("|")} | (LOCAL_TYPES if exp == "LocalBusiness" else set())
                if have & opts or n_schema >= 300:
                    continue
                n_schema += 1
                first = exp.split("|")[0]
                h1 = next((x["text"] for x in p["headings"] if x["level"] == 1 and x["text"]), p.get("title") or "")
                if first == "BreadcrumbList":
                    segs = [x for x in urlparse(p["url"]).path.split("/") if x]
                    items = [{"@type": "ListItem", "position": 1, "name": "Home", "item": dom + "/"}]
                    for k, seg in enumerate(segs, 2):
                        items.append({"@type": "ListItem", "position": k,
                                      "name": h1 if k == len(segs) + 1 else seg.replace("-", " ").title(),
                                      "item": dom + "/" + "/".join(segs[:k - 1]) + "/"})
                    snip = {"@context": "https://schema.org", "@type": "BreadcrumbList", "itemListElement": items}
                elif first in ("Article",):
                    snip = {"@context": "https://schema.org", "@type": "Article", "headline": h1, "url": p["url"],
                            "image": p.get("og_image") or "<image URL>", "datePublished": "<YYYY-MM-DD>",
                            "author": {"@type": "Person", "name": "<Author name>"},
                            "publisher": {"@id": f"{dom}/#organization"}}
                elif first == "Product":
                    snip = {"@context": "https://schema.org", "@type": "Product", "name": h1, "image": p.get("og_image") or "<image>",
                            "offers": {"@type": "Offer", "price": "<price>", "priceCurrency": "INR",
                                       "availability": "https://schema.org/InStock", "url": p["url"]}}
                elif first == "LocalBusiness":
                    snip = {"@context": "https://schema.org", "@type": "LocalBusiness", "name": brand, "url": p["url"],
                            "telephone": tel or "<phone>", "address": {"@type": "PostalAddress", "streetAddress": "<street>",
                                                                       "addressLocality": "<city>", "postalCode": "<PIN>", "addressCountry": "IN"},
                            "parentOrganization": {"@id": f"{dom}/#organization"}}
                elif first == "Service":
                    snip = {"@context": "https://schema.org", "@type": "Service", "name": h1, "url": p["url"],
                            "provider": {"@id": f"{dom}/#organization"}, "areaServed": "<city/region>"}
                elif first in ("CollectionPage", "ContactPage", "AboutPage", "Organization"):
                    snip = {"@context": "https://schema.org", "@type": first, "name": h1, "url": p["url"],
                            "isPartOf": {"@id": f"{dom}/#website"}}
                else:
                    snip = {"@context": "https://schema.org", "@type": first, "name": h1, "url": p["url"]}
                add(f"Missing {first} schema", "Medium", p["url"], "<head> JSON-LD", f"Has: {', '.join(sorted(page_types(p))) or 'none'}",
                    f"Add {first} (fill <placeholders>)", where_for("Schema type expected", fam), jsonld(snip), "Developer")
        # expert fixes from Claude
        for it in self.a.get("issues", []):
            for f in it.get("fixes", []):
                add(it.get("issue", ""), it.get("severity", "Medium"), f.get("page", ""), f.get("element", ""), f.get("current", ""),
                    f.get("change_to", ""), f.get("where", it.get("where", "")), f.get("code", ""), it.get("owner", ""))
        if len(rows) >= cap:
            rows.append(["(Fix list truncated)", "Info", "", "", "", f"Showing first {cap} fixes; see other tabs for the rest", "", "", ""])
        order = {s: i for i, s in enumerate(SEV_ORDER)}
        rows.sort(key=lambda r: order.get(r[1], 9))
        return [r for r in rows if in_scope({"area": self._area_of(r[0]), "category": self._cat_of(r[0]), "issue": r[0]}, self.scope)]

    def _area_of(self, label):
        l = label.lower()
        if any(k in l for k in ("schema", "rating", "organization")):
            return "Schema"
        if any(k in l for k in ("title", "meta desc", "h1", "heading", "image", "alt")):
            return "On-Page"
        return next((i["area"] for i in self.issues if i["issue"] == label), "Technical")

    def _cat_of(self, label):
        l = label.lower()
        if "image" in l or "alt" in l:
            return "Images"
        if "sitemap" in l:
            return "Sitemap"
        if "redirect" in l:
            return "Redirects"
        if "broken" in l:
            return "URL errors"
        if "noindex" in l or "canonical" in l:
            return "Indexability"
        if "meta desc" in l:
            return "Meta description"
        return next((i["category"] for i in self.issues if i["issue"] == label), "")

    def run(self):
        self.technical()
        self.cwv()
        self.onpage()
        self.schema()
        self.offpage()
        for it in self.a.get("issues", []):
            urls = it.get("urls") or []
            self.issues.append({"area": it.get("area", "Content"), "category": it.get("category", "Expert review"),
                                "issue": it.get("issue", ""), "severity": it.get("severity", "Medium"),
                                "count": len(urls) or 1, "urls": urls, "evidence": it.get("evidence", ""),
                                "why": it.get("why", ""), "fix": it.get("fix", ""), "owner": it.get("owner", ""),
                                "where": it.get("where", ""), "code": it.get("code", ""),
                                "source": "Claude (claude-seo)"})
            for u in urls:
                self.page_issues[u].append(it.get("issue", ""))
        # verification from Claude (instead of duplicate "confirmed" issues)
        for v in self.a.get("verify", []):
            title = (v.get("issue") or "").strip().lower()
            for i in self.issues:
                if i["source"] == "Crawler" and title and (i["issue"].lower() == title or i["issue"].lower().startswith(title[:40])):
                    verdict = (v.get("verdict") or "").strip()
                    i["verified"] = f"{verdict}: {v.get('note', '')}".strip(": ")
                    if verdict.lower().startswith("false"):
                        i["severity"] = "Info"
                        i["issue"] = "FALSE POSITIVE - " + i["issue"]
        for i in self.issues:
            i.setdefault("verified", "")
            if not i.get("where"):
                i["where"] = where_for(i["issue"].replace("FALSE POSITIVE - ", ""), self.fam)
        if self.scope:
            self.issues = [i for i in self.issues if in_scope(i, self.scope)]
            keep = {u for i in self.issues for u in i["urls"]}
            self.page_issues = defaultdict(list, {u: [x for x in v if any(x == i["issue"] for i in self.issues)]
                                                  for u, v in self.page_issues.items() if u in keep})
        self.issues.sort(key=lambda i: (SEV_ORDER.index(i["severity"]) if i["severity"] in SEV_ORDER else 9,
                                        i["area"], -i["count"]))
        return self


# ============================================================================ sheets
def build_tables(au):
    c, P = au.c, au.pages
    m = c["meta"]
    T = {}  # name -> (headers, rows, widths)

    # Issue tracker
    rows = []
    for n, i in enumerate(au.issues, 1):
        ex = "\n".join(i["urls"][:5]) + (f"\n… +{len(i['urls']) - 5} more" if len(i["urls"]) > 5 else "")
        rows.append([n, i["area"], i["category"], i["issue"], i["severity"], i["count"],
                     (i["evidence"] + ("\n" if i["evidence"] and ex else "") + ex).strip(),
                     i["why"], i["fix"], i.get("where", ""), i["owner"], "Open", i.get("verified", ""), "", i["source"]])
    T["Issue Tracker"] = (["#", "Area", "Category", "Issue", "Severity", "Affected URLs", "Evidence / example URLs",
                           "Why it matters", "How to fix", f"Where to fix ({au.platform})", "Owner", "Status",
                           "Verified by Claude", "Notes", "Found by"],
                          rows, [5, 11, 16, 40, 10, 9, 55, 40, 50, 55, 14, 11, 30, 25, 14])
    fx = au.fix_rows()
    T["Fix List"] = (["#", "Issue", "Severity", "Page / file to edit", "Element / location", "Current", "Change to",
                      f"Where to implement ({au.platform})", "Code / snippet", "Owner", "Status"],
                     [[k] + r + ["Open"] for k, r in enumerate(fx, 1)],
                     [5, 28, 10, 55, 40, 45, 55, 60, 60, 14, 11])

    # Page inventory
    rows = []
    for p in P:
        h1 = [h["text"] for h in p["headings"] if h["level"] == 1]
        imgs = {i["url"]: i for i in c.get("images", [])}
        big = sum(1 for im in p["images"] if (imgs.get(im["src"], {}).get("bytes") or 0) > 102400)
        idx = "No (noindex)" if "noindex" in ((p.get("meta_robots") or "") + (p.get("x_robots_tag") or "")).lower() else "Yes"
        rows.append([p["url"], p["page_type"], p["status"], idx, "Yes" if p["url"] in au.sitemap else "No",
                     au.inlink_counts.get(p["url"], 0), p.get("title") or "", len(p.get("title") or ""),
                     len(p.get("meta_description") or ""), len(h1), " | ".join(h1)[:150], p["word_count"],
                     ", ".join(sorted(page_types(p))), len(p["images"]), big, p.get("elapsed_ms"),
                     len(au.page_issues.get(p["url"], [])), "; ".join(au.page_issues.get(p["url"], []))[:500],
                     au.a.get("page_notes", {}).get(p["url"], "")])
    # keep pages with page-specific issues; site-wide template issues are listed once in the Issue Tracker
    n_pages = len(P)
    sitewide = {i["issue"] for i in au.issues if n_pages > 3 and i["count"] >= 0.8 * n_pages}
    keep = []
    for r in rows:
        own = [x for x in au.page_issues.get(r[0], []) if x not in sitewide]
        if own:
            r[16], r[17] = len(own), "; ".join(own)[:500]
            keep.append(r)
    rows = keep
    T["Page Inventory"] = (["URL", "Page type", "Status", "Indexable", "In sitemap", "Internal inlinks", "Title",
                            "Title length", "Meta desc length", "H1 count", "H1 text", "Word count", "Schema types",
                            "Images", "Images >100KB", "Response ms", "Issue count", "Issues", "Expert note"],
                           rows, [55, 22, 7, 12, 9, 9, 45, 8, 8, 7, 35, 8, 30, 7, 9, 9, 8, 70, 45])

    # Redirects
    rows = []
    for s in au.status.values():
        if not s.get("chain"):
            continue
        chain = " → ".join([h["url"] for h in s["chain"]] + [s.get("final_url") or "?"])
        hops = len(s["chain"])
        fi = ", ".join(au.found_in.get(s["url"], []))
        src = au.inlinks.get(s["url"], [])
        fix = ("Remove from sitemap / replace with final URL. " if s["url"] in au.sitemap else "") + \
              ("Collapse chain into one 301 to the final URL. " if hops > 1 else "") + \
              ("Update internal links to the final URL." if src else "")
        if not fix.strip():
            continue  # legacy 301 that nobody links to and is not in the sitemap is fine
        rows.append([s["url"], s["chain"][0]["status"], s["chain"][0]["to"], s.get("final_url") or "",
                     s.get("final_status"), hops, chain, fi, len(src), "\n".join(src[:5]), fix.strip()])
    T["Tech - Redirects"] = (["Source URL", "Redirect code", "Redirects to (next hop)", "Final URL", "Final status",
                              "Hops", "Full chain", "Found in", "Linked from (count)", "Linked from (examples)", "Fix"],
                             rows, [50, 9, 50, 50, 8, 6, 80, 22, 9, 55, 50])

    # URL errors
    rows = []
    for s in au.status.values():
        fs = s.get("final_status") or 0
        if is_broken(s):
            kind = (f"{fs} " + ("Not Found" if fs == 404 else "Gone" if fs == 410 else "Forbidden" if fs == 403
                                else "Server error" if fs >= 500 else "Client error")) if fs else (s.get("error") or "")
            src = au.inlinks.get(s["url"], [])
            rows.append([s["url"], fs or "", kind, ", ".join(au.found_in.get(s["url"], [])),
                         "Yes" if s["url"] in au.sitemap else "No", len(src), "\n".join(src[:8]),
                         "Remove/replace the link on the source pages, or 301 this URL to the closest live page."
                         + (" Remove from sitemap." if s["url"] in au.sitemap else "")])
    for s in au.status.values():
        if s.get("unreachable"):
            rows.append([s["url"], "", "NOT CONFIRMED - no response during crawl (server throttling)",
                         ", ".join(au.found_in.get(s["url"], [])), "Yes" if s["url"] in au.sitemap else "No",
                         len(au.inlinks.get(s["url"], [])), "", "Open in a browser to confirm before fixing anything."])
    for s in au.status.values():
        if s.get("error") and "robots" in s["error"]:
            rows.append([s["url"], "", "Blocked by robots.txt (not crawled)", ", ".join(au.found_in.get(s["url"], [])),
                         "Yes" if s["url"] in au.sitemap else "No", len(au.inlinks.get(s["url"], [])), "",
                         "Intended? If it should rank, allow it in robots.txt."])
    T["Tech - URL Errors"] = (["URL", "Status", "Error type", "Found in", "In sitemap", "Linked from (count)",
                               "Linked from (pages)", "Fix"], rows, [60, 7, 26, 22, 9, 9, 60, 55])

    # Indexability
    rows = []
    for p in P:
        robots = ((p.get("meta_robots") or "") + " " + (p.get("x_robots_tag") or "")).strip()
        c_url = p.get("canonical") or ""
        cst = au.status.get(c_url, {}).get("final_status") if c_url else ""
        self_c = "Self" if c_url and c_url.rstrip("/") == p["url"].rstrip("/") else ("Missing" if not c_url else "Other URL")
        issue = []
        if "noindex" in robots.lower():
            issue.append("noindex" + (" + in sitemap (conflict)" if p["url"] in au.sitemap else ""))
        if self_c == "Missing":
            issue.append("no canonical")
        if self_c == "Other URL":
            issue.append("canonicalised elsewhere")
        if not issue:
            continue
        rows.append([p["url"], p["page_type"], robots or "index, follow (default)", c_url, self_c, cst or "",
                     "Yes" if p["url"] in au.sitemap else "No", ", ".join(h["lang"] or "" for h in p.get("hreflang", [])),
                     "; ".join(issue)])
    T["Tech - Indexability"] = (["URL", "Page type", "Robots directives", "Canonical", "Canonical type",
                                 "Canonical status", "In sitemap", "hreflang", "Result"],
                                rows, [55, 22, 25, 55, 12, 9, 9, 15, 40])

    # Site checks
    sc = c.get("site_checks", {})
    def chk(name, ok, detail, warn=False):
        return [name, "Pass" if ok else ("Warn" if warn else "Fail"), detail]
    rows = []
    if m["mode"] != "single page":
        r = sc.get("robots_txt", {})
        rows.append(chk("robots.txt present", r.get("found"), f"{r.get('url')} -> HTTP {r.get('status')}"))
        sm_count = sum(s.get("count", 0) for s in c.get("sitemaps", []) if s.get("type") == "urlset")
        rows.append(chk("XML sitemap found", bool(c.get("sitemaps")),
                        f"{len(c.get('sitemaps', []))} sitemap file(s), {sm_count} URL entries, {len(au.sitemap)} unique"))
        bad_sm = [s for s in c.get("sitemaps", []) if str(s.get("type", "")).startswith(("invalid", "not a sitemap", "error"))]
        warn_sm = [s for s in c.get("sitemaps", []) if s.get("error") and s not in bad_sm]
        rows.append(chk("Sitemap files valid", not bad_sm,
                        "; ".join(f"{s['url']}: {s['type']} ({s.get('error', '')})" for s in bad_sm) or "All parse OK"))
        if warn_sm:
            rows.append(["Sitemap formatting", "Warn", "; ".join(f"{s['url']}: {s['error']}" for s in warn_sm)])
        h = sc.get("http_root", {})
        if m["start_url"].startswith("https://"):
            rows.append(chk("HTTP -> HTTPS redirect", str(h.get("final_url") or "").startswith("https://"),
                            f"{h.get('url')} -> {h.get('final_url')} ({h.get('hops')} hop)"))
        else:
            rows.append(chk("Site served over HTTPS", False, f"Start URL is http:// ({m['start_url']}) - move to HTTPS"))
        a = sc.get("alt_www_root", {})
        if a.get("error"):
            rows.append(["www / non-www consolidated", "Warn", f"{a.get('url')} not reachable ({a.get('error')[:80]}) - check DNS/SSL for the other host"])
        else:
            rows.append(chk("www / non-www consolidated", not (a.get("final_status") == 200 and a.get("hops") == 0),
                            f"{a.get('url')} -> {a.get('final_url')} ({a.get('hops')} hop)"))
        rows.append(chk("Missing pages return 404", sc.get("soft_404_test", {}).get("status") in (404, 410),
                        f"Random URL returned {sc.get('soft_404_test', {}).get('status')}"))
    ttfb = sorted(p.get("elapsed_ms") or 0 for p in P)
    if ttfb:
        med = ttfb[len(ttfb) // 2]
        rows.append(chk("Median HTML response time", med < 800, f"{med} ms (target < 800 ms)", warn=med < 1500))
    rows.append(chk("Mixed content", not any(p.get("mixed_content") for p in P),
                    f"{sum(1 for p in P if p.get('mixed_content'))} page(s) load http:// resources"))
    rows.append(chk("Viewport meta on all pages", all(p.get("viewport") for p in P),
                    f"{sum(1 for p in P if not p.get('viewport'))} page(s) missing"))
    rows.append(chk("HTML lang on all pages", all(p.get("lang") for p in P),
                    f"{sum(1 for p in P if not p.get('lang'))} page(s) missing"))
    cc = Counter((p.get("cache_control") or "none").split(",")[0] for p in P)
    rows.append(chk("HTML cache headers", not any("no-store" in (p.get("cache_control") or "") for p in P),
                    "; ".join(f"{k}: {v}" for k, v in cc.most_common(3)), warn=True))
    T["Tech - Site Checks"] = (["Check", "Result", "Detail"], [r for r in rows if r[1] != "Pass"], [35, 9, 100])

    # Sitemap URLs
    rows = []
    for s in c.get("sitemap_urls", []):
        st = au.status.get(s["url"], {})
        res = "OK" if st.get("final_status") == 200 and not st.get("chain") else \
            ("Redirects" if st.get("chain") else ("Error" if (st.get("final_status") or 0) >= 400 else
             ("Not confirmed (no response)" if st.get("unreachable") else "Not checked")))
        if res in ("OK", "Not checked") and s.get("listed", 1) == 1:
            continue
        if res == "OK":
            res = f"Listed {s.get('listed')} times"
        rows.append([s["url"], st.get("status") or "", st.get("final_url") or "", st.get("final_status") or "",
                     ", ".join(s["sitemaps"]), s.get("listed", 1), res])
    T["Tech - Sitemap URLs"] = (["URL", "First status", "Final URL", "Final status", "Sitemap file(s)", "Times listed",
                                 "Result"], rows, [60, 9, 55, 9, 45, 8, 12])

    # Core Web Vitals
    rows = []
    for x in c.get("pagespeed") or []:
        f, o = x.get("field_url") or {}, x.get("field_origin") or {}
        if not x.get("error") and (x.get("perf_score") or 0) >= 90 and \
                all(f.get(k) in (None, "FAST") for k in ("lcp_cat", "inp_cat", "cls_cat")):
            continue
        fmt = lambda v, cat: f"{v} ({cat})" if v is not None else "no field data"
        rows.append([x["url"], x.get("perf_score", "") if not x.get("error") else x["error"],
                     round(x["lab_lcp_ms"]) if x.get("lab_lcp_ms") else "", round(x["lab_tbt_ms"]) if x.get("lab_tbt_ms") else "",
                     round(x["lab_cls"], 3) if x.get("lab_cls") is not None else "",
                     fmt(f.get("lcp_ms"), f.get("lcp_cat")), fmt(f.get("inp_ms"), f.get("inp_cat")),
                     fmt(f.get("cls"), f.get("cls_cat")), f"{o.get('overall') or 'n/a'} (LCP {o.get('lcp_ms')} / INP {o.get('inp_ms')})",
                     "\n".join(x.get("top_opportunities") or [])])
    T["Tech - Core Web Vitals"] = (["URL", "Mobile perf score (lab)", "Lab LCP ms", "Lab TBT ms", "Lab CLS",
                                    "Real-user LCP (p75)", "Real-user INP (p75)", "Real-user CLS (p75)",
                                    "Whole-site real-user", "Top fixes (est. saving)"],
                                   rows, [50, 12, 10, 10, 8, 18, 18, 18, 30, 60])

    # Titles & meta
    tcount = Counter((p.get("title") or "").strip().lower() for p in P if p.get("title"))
    mcount = Counter((p.get("meta_description") or "").strip().lower() for p in P if p.get("meta_description"))
    rows = []
    for p in P:
        t, d = p.get("title") or "", p.get("meta_description") or ""
        ti = "Missing" if not t else ">60 chars" if len(t) > 60 else "<30 chars" if len(t) < 30 else "OK"
        if t and tcount[t.strip().lower()] > 1:
            ti += " + duplicate"
        di = "Missing" if not d else ">160 chars" if len(d) > 160 else "<70 chars" if len(d) < 70 else "OK"
        if d and mcount[d.strip().lower()] > 1:
            di += " + duplicate"
        if ti == "OK" and di == "OK":
            continue
        rows.append([p["url"], p["page_type"], t, len(t), ti, d, len(d), di, p.get("og_title") or "",
                     "Yes" if p.get("og_image") else "No"])
    T["OnPage - Titles & Meta"] = (["URL", "Page type", "Title", "Title len", "Title check", "Meta description",
                                    "Desc len", "Desc check", "OG title", "OG image"],
                                   rows, [50, 22, 50, 8, 16, 60, 8, 16, 35, 8])

    # Headings
    rows = []
    for p in P:
        hs = p["headings"]
        content = [h for h in hs if h["region"] in ("main", "body")]
        issues = []
        h1n = sum(1 for h in hs if h["level"] == 1)
        if h1n == 0:
            issues.append("No H1")
        if h1n > 1:
            issues.append(f"{h1n} H1 tags")
        prev = 0
        for h in content:
            if prev and h["level"] > prev + 1:
                issues.append(f"Skips H{prev} -> H{h['level']} at '{h['text'][:40]}'")
            prev = h["level"]
        if content and content[0]["level"] != 1:
            issues.append(f"Content starts with H{content[0]['level']}")
        if any(h["empty"] for h in hs):
            issues.append("Empty heading")
        if not issues:
            continue  # header/footer headings are reported once, in the Header-Footer tab
        seq = " > ".join(f"H{h['level']}" for h in content)
        outline = "\n".join(("  " * (h["level"] - 1)) + f"H{h['level']}: {h['text'] or '(empty)'}" for h in content)[:3000]
        rows.append([p["url"], p["page_type"], h1n, sum(1 for h in hs if h["level"] == 2), seq[:300],
                     "; ".join(issues), outline])
    T["OnPage - Headings"] = (["URL", "Page type", "H1 count", "H2 count", "Heading order (content)", "Hierarchy issues",
                               "Content outline"], rows, [50, 22, 7, 7, 45, 55, 80])

    # Header / footer headings (template level, de-duplicated)
    groups = defaultdict(set)
    for p in P:
        for h in p["headings"]:
            if h["region"] in ("header", "footer", "nav", "aside"):
                groups[(h["region"], h["level"], h["text"] or "(empty)")].add(p["url"])
    total = len(P) or 1
    rows = []
    for (region, lvl, text), urls in sorted(groups.items(), key=lambda kv: -len(kv[1])):
        scope = "Site-wide template" if len(urls) / total >= 0.6 and len(urls) > 1 else "Some pages"
        rows.append([region.title(), f"H{lvl}", text, len(urls), scope, "\n".join(sorted(urls)[:5]),
                     f"Change this H{lvl} to a <p> or <div> with a CSS class (Elementor: Heading widget > Advanced/"
                     f"HTML Tag = div). Visual style stays the same."])
    T["OnPage - Header-Footer H"] = (["Region", "Tag", "Heading text", "Pages affected", "Scope", "Example pages",
                                      "How to fix"], rows, [10, 6, 40, 9, 18, 55, 70])

    # Images
    rows = []
    alt_map = defaultdict(set)
    for p in P:
        for im in p["images"]:
            alt_map[im["src"]].add((im.get("alt") or "").strip() if im["has_alt_attr"] else None)
    for i in sorted(c.get("images", []), key=lambda x: -(x.get("bytes") or 0)):
        kb = round((i.get("bytes") or 0) / 1024, 1) if i.get("bytes") else ""
        ext = os.path.splitext(urlparse(i["url"]).path)[1].lower().lstrip(".") or (i.get("type") or "").split("/")[-1]
        alts = alt_map.get(i["url"], set())
        alt_txt = "; ".join(sorted(a for a in alts if a))
        missing = "Yes" if (None in alts or "" in alts) else "No"
        fix = []
        if kb and kb > 100:
            fix.append(f"Compress to <100 KB" + (" and convert to WebP" if ext in ("jpg", "jpeg", "png", "gif", "bmp") else ""))
        if missing == "Yes":
            fix.append("Add alt text")
        if (i.get("status") or 0) >= 400:
            fix.append("Broken - re-upload or remove")
        if not fix:
            continue
        rows.append([i["url"], kb, ext, i.get("status") or i.get("error") or "",
                     "Yes" if kb and kb > 100 else ("Not checked" if kb == "" else "No"), missing, alt_txt[:150],
                     i.get("page_count", 0), "\n".join(i.get("pages", [])[:5]), "; ".join(fix)])
    T["OnPage - Images"] = (["Image URL", "Size KB", "Format", "Status", "Over 100 KB", "Missing alt", "Alt text",
                             "Used on (pages)", "Example pages", "Fix"], rows, [60, 9, 7, 7, 10, 9, 35, 8, 50, 40])

    # Schema per page
    rows = []
    for p in P:
        have = page_types(p)
        hl = {t.lower() for t in have}
        missing = []
        for exp in EXPECTED_SCHEMA.get(p["page_type"], []):
            opts = {o.lower() for o in exp.split("|")}
            if exp == "LocalBusiness":
                opts |= LOCAL_TYPES
            if not (hl & opts):
                missing.append(exp.replace("|", " or "))
        notes = []
        for n in p["schema"]:
            tl = {t.lower() for t in types_of(n)}
            if "product" in tl:
                off = n.get("offers") or {}
                off = off[0] if isinstance(off, list) and off else off
                for k in ("price", "priceCurrency", "availability"):
                    if not (isinstance(off, dict) and off.get(k)):
                        notes.append(f"Product.offers missing {k}")
                if not n.get("image"):
                    notes.append("Product missing image")
            if tl & {"article", "blogposting", "newsarticle"}:
                for k in ("headline", "author", "datePublished", "image"):
                    if not n.get(k):
                        notes.append(f"Article missing {k}")
        if not missing and not p.get("schema_errors") and not notes:
            continue
        rows.append([p["url"], p["page_type"], ", ".join(sorted(have)) or "None",
                     ", ".join(EXPECTED_SCHEMA.get(p["page_type"], [])).replace("|", "/"),
                     ", ".join(missing) or "None", "; ".join(p.get("schema_errors", []) + notes) or "-"])
    T["Schema - Per Page"] = (["URL", "Page type", "Schema types found", "Expected for page type", "Missing",
                               "Errors / missing properties"], rows, [55, 22, 40, 40, 30, 55])

    # Organization & Local
    rows = []
    org_nodes, loc_nodes = [], []
    for p in P:
        for n in p["schema"]:
            tl = {t.lower() for t in types_of(n)}
            if tl & (ORG_TYPES - LOCAL_TYPES):
                org_nodes.append((p["url"], n))
            if tl & LOCAL_TYPES:
                loc_nodes.append((p["url"], n))
    org_req = [("name", True), ("url", True), ("logo", True), ("sameAs", False), ("contactPoint", False),
               ("@id", False), ("address", False)]
    if not org_nodes and loc_nodes:
        org_nodes = loc_nodes  # LocalBusiness is an Organization subtype
    if org_nodes:
        u, n = org_nodes[0]
        rows.append(["Organization", "Present", "Pass", f"Found on {len({x[0] for x in org_nodes})} of {len(P)} pages", ""])
        for k, req in org_req:
            v = n.get(k)
            rows.append(["Organization", f"{k} ({'required' if req else 'recommended'})",
                         "Pass" if v else ("Fail" if req else "Warn"), short(json.dumps(v, ensure_ascii=False), 150) if v else "missing",
                         "" if v else f"Add '{k}' to the Organization node."])
        names = {str(x[1].get("name")) for x in org_nodes}
        rows.append(["Organization", "Consistent name across pages", "Pass" if len(names) == 1 else "Fail",
                     "; ".join(sorted(names)), "" if len(names) == 1 else "Use one Organization node with fixed @id."])
    else:
        rows.append(["Organization", "Present", "Fail", "Not found on any crawled page",
                     "Add Organization schema site-wide (Yoast/Rank Math: Knowledge Graph settings)."])
    loc_req = [("name", True), ("address", True), ("telephone", True), ("url", False), ("image", False),
               ("geo", False), ("openingHoursSpecification|openingHours", False), ("priceRange", False),
               ("areaServed", False)]
    if loc_nodes:
        rows.append(["LocalBusiness", "Present", "Pass",
                     f"Found on {len({x[0] for x in loc_nodes})} page(s): types " +
                     ", ".join(sorted({t for _, n in loc_nodes for t in types_of(n)})), ""])
        for u, n in loc_nodes[:25]:
            for k, req in loc_req:
                v = next((n.get(x) for x in k.split("|") if n.get(x)), None)
                if k == "address" and isinstance(v, dict):
                    miss = [f for f in ("streetAddress", "addressLocality", "postalCode", "addressCountry") if not v.get(f)]
                    rows.append([f"LocalBusiness ({short(u, 60)})", "address fields", "Pass" if not miss else "Warn",
                                 short(json.dumps(v, ensure_ascii=False), 150), ("Add " + ", ".join(miss)) if miss else ""])
                    continue
                rows.append([f"LocalBusiness ({short(u, 60)})", f"{k} ({'required' if req else 'recommended'})",
                             "Pass" if v else ("Fail" if req else "Warn"),
                             short(json.dumps(v, ensure_ascii=False), 150) if v else "missing",
                             "" if v else f"Add '{k.split('|')[0]}' (must match Google Business Profile)."])
            if n.get("aggregateRating") or n.get("review"):
                rows.append([f"LocalBusiness ({short(u, 60)})", "Self-serving reviews", "Fail", "aggregateRating/review present",
                             "Remove unless genuine on-page reviews."])
    else:
        rows.append(["LocalBusiness", "Present", "Warn", "Not found on any crawled page",
                     "If the business serves customers locally, add LocalBusiness (or a subtype like Dentist) on "
                     "home/contact/location pages with NAP matching Google Business Profile."])
    T["Schema - Org & Local"] = (["Entity", "Check", "Result", "Value found", "Fix"],
                                 [r for r in rows if r[2] != "Pass"], [35, 35, 8, 70, 60])

    # Off-page
    rows = []
    for e in sorted(c.get("external_links", []), key=lambda x: (-(x.get("status") or 999) if (x.get("status") or 0) >= 400 else 0)):
        rels = set()
        for p in P:
            for ln in p["links"]:
                if ln["href"] == e["url"]:
                    rels.add(ln["rel"] or "follow")
        dom = urlparse(e["url"]).netloc
        kind = "Social profile" if any(s in dom for s in SOCIAL) else "Outbound link"
        st = e.get("status")
        res = "Broken" if (st or 0) >= 400 or (e.get("error") and st is None) else ("Not checked" if st is None else "OK")
        if res != "Broken":
            continue
        rows.append([kind, e["url"], dom, st or e.get("error") or "", res, ", ".join(sorted(rels)),
                     e.get("page_count", 0), "\n".join(e.get("pages", [])[:5])])
    off = au.a.get("offpage", {})
    for rd in off.get("referring_domains", []):
        rows.append(["Referring domain (backlink)", rd.get("source", ""), rd.get("domain", ""), "", "", "", "", rd.get("note", "")])
    for note in off.get("notes", []) + ([off["backlinks_note"]] if off.get("backlinks_note") else []):
        rows.append(["Note", "", "", "", "", "", "", note])
    if not off:
        rows.append(["Note", "", "", "", "", "", "",
                     "Backlink profile needs Search Console (Links report) or a backlink tool; this tab covers outbound "
                     "links, social profiles and NAP signals visible on the site."])
    T["OffPage - Links & Profiles"] = (["Type", "URL", "Domain", "Status", "Result", "rel", "Linked from (count)",
                                        "Linked from / note"], rows, [22, 60, 25, 8, 10, 14, 9, 60])

    # Page types
    by = defaultdict(list)
    for p in P:
        by[p["page_type"]].append(p)
    rows = []
    for t, ps in sorted(by.items(), key=lambda kv: -len(kv[1])):
        iss = Counter(i for p in ps for i in au.page_issues.get(p["url"], []))
        rows.append([t, len(ps), round(sum(p["word_count"] for p in ps) / len(ps)),
                     "; ".join(f"{k} ({v})" for k, v in iss.most_common(4)) or "None",
                     ", ".join(EXPECTED_SCHEMA.get(t, [])).replace("|", "/"), "\n".join(p["url"] for p in ps[:5])])
    T["Page Types"] = (["Page type", "Pages", "Avg words", "Top issues (pages)", "Expected schema", "Example URLs"],
                       rows, [30, 7, 9, 80, 35, 60])

    # Action plan
    rows = [[a.get("phase", ""), a.get("task", ""), a.get("owner", ""), "Open"] for a in au.a.get("action_plan", [])]
    if not rows:
        phase = {"Critical": "Week 1", "High": "Weeks 1-2", "Medium": "Weeks 3-4", "Low": "Month 2", "Info": "Review"}
        rows = [[phase.get(i["severity"], ""), f"{i['issue']} ({i['count']} URLs) - {i['fix']}", i["owner"], "Open"]
                for i in au.issues if i["severity"] != "Info"]
    T["Action Plan"] = (["Phase", "Task", "Owner", "Status"], rows, [14, 110, 16, 11])

    # Data sources
    rows = [["Crawler (crawl_site.py)", m["mode"], f"{m['pages_parsed']} HTML pages parsed; {m['urls_status_checked']} URLs status-checked; "
             f"{m['images_checked']}/{m['images_found']} images sized; {m['external_links']} external links; "
             f"{m.get('unreachable', 0)} URLs gave no response (not counted as broken)", m["crawled_at"],
             "; ".join(m.get("notes", [])) or "None"],
            ["Sitemaps", "robots.txt + common paths", f"{len(c.get('sitemaps', []))} files, {len(au.sitemap)} unique URLs", m["crawled_at"], ""],
            ["claude-seo skill (expert review)", au.a.get("audit_label", ""), "Content quality, E-E-A-T, strategy, extra issues",
             m["crawled_at"], "Judgement-based; verify before client delivery"],
            ["Not included", "", "Search Console, GA4, backlink databases, rank tracking", "", "Connect GSC for real index/backlink data"]]
    T["Data Sources"] = (["Source", "Scope", "Coverage", "Run at", "Limitations"], rows, [32, 25, 70, 20, 60])
    return T


# ============================================================================ write
def write_xlsx(au, T, path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import CellIsRule, FormulaRule

    m = au.c["meta"]
    wb = Workbook()
    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")

    # ---------- Summary
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = f"SEO Audit: {m['domain']}"
    ws["A1"].font = Font(bold=True, size=16, color="1F3864")
    ws["A2"] = (f"{au.a.get('audit_label', m['mode'].title())} | Start URL: {m['start_url']} | "
                f"Platform: {au.platform} | Crawled: {m['crawled_at']}")
    ws["A2"].font = Font(italic=True, color="595959")
    r = 4
    if au.a.get("executive_summary"):
        ws.cell(r, 1, "Executive summary").font = Font(bold=True, size=12)
        r += 1
        ws.cell(r, 1, au.a["executive_summary"]).alignment = wrap
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 90
        r += 2
    ws.cell(r, 1, "Crawl coverage").font = Font(bold=True, size=12)
    r += 1
    for k, v in (("HTML pages analysed", m["pages_parsed"]), ("URLs status-checked", m["urls_status_checked"]),
                 ("Unique sitemap URLs", len(au.sitemap)), ("Images sized", f"{m['images_checked']} / {m['images_found']}"),
                 ("External links", m["external_links"]), ("Crawl notes", "; ".join(m.get("notes", [])) or "None")):
        ws.cell(r, 1, k)
        ws.cell(r, 2, v).alignment = Alignment(horizontal="left")
        r += 1
    r += 1
    scores = au.a.get("scores") or {}
    if scores:
        ws.cell(r, 1, "Scores (claude-seo, 0-100)").font = Font(bold=True, size=12)
        r += 1
        for k, v in scores.items():
            ws.cell(r, 1, k)
            ws.cell(r, 2, v)
            r += 1
        r += 1
    ws.cell(r, 1, "Issues by severity").font = Font(bold=True, size=12)
    r += 1
    for j, h in enumerate(["Severity", "Issues", "Done", "Remaining"], 1):
        cell = ws.cell(r, j, h)
        cell.fill, cell.font = hdr_fill, hdr_font
    r += 1
    for sev in SEV_ORDER:
        ws.cell(r, 1, sev).fill = PatternFill("solid", fgColor=SEV_FILL[sev])
        if sev == "Critical":
            ws.cell(r, 1).font = Font(color="FFFFFF", bold=True)
        ws.cell(r, 2, f"=COUNTIFS('Issue Tracker'!$E:$E,$A{r})")
        ws.cell(r, 3, f"=COUNTIFS('Issue Tracker'!$E:$E,$A{r},'Issue Tracker'!$L:$L,\"Done\")")
        ws.cell(r, 4, f"=B{r}-C{r}")
        r += 1
    r += 1
    ws.cell(r, 1, "Issues by area").font = Font(bold=True, size=12)
    r += 1
    for j, h in enumerate(["Area", "Issues"], 1):
        cell = ws.cell(r, j, h)
        cell.fill, cell.font = hdr_fill, hdr_font
    r += 1
    for area in sorted({i["area"] for i in au.issues}):
        ws.cell(r, 1, area)
        ws.cell(r, 2, f"=COUNTIFS('Issue Tracker'!$B:$B,$A{r})")
        r += 1
    r += 1
    ws.cell(r, 1, "How to use this workbook").font = Font(bold=True, size=12)
    r += 1
    guide = [("Fix List  <- start here", "One row per fix: which page/file, current value, what to change it to, "
                                         "where to change it on this platform, and copy-paste code."),
             ("Issue Tracker", "Every issue with severity, fix and owner. Update 'Status' as you work."),
             ("Page Inventory", "Only pages that have at least one issue: type, title/meta/H1, schema and the issue list."),
             ("Only problems", "Every tab lists only items that need fixing. A missing tab = nothing to fix in that area."),
             ("Tech - *", "Redirects (which URL goes where), URL errors with linking pages, indexability, site checks, sitemap URLs."),
             ("OnPage - *", "Titles & meta, heading hierarchy, headings inside header/footer, images (>100 KB, alt)."),
             ("Schema - *", "Schema per page vs what each page type should have; Organization & LocalBusiness checks."),
             ("OffPage - *", "Outbound links, social profiles vs sameAs, NAP, backlink notes."),
             ("Page Types", "Pages grouped by type (landing, product, blog...) with their most common issues.")]
    for k, v in guide:
        ws.cell(r, 1, k).font = Font(bold=True)
        ws.cell(r, 2, v)
        r += 1
    ws.column_dimensions["A"].width = 32
    for col in "BCDE":
        ws.column_dimensions[col].width = 22

    # ---------- data sheets
    order = ["Fix List", "Issue Tracker", "Page Inventory", "Page Types", "Tech - Redirects", "Tech - URL Errors",
             "Tech - Indexability", "Tech - Sitemap URLs", "Tech - Site Checks", "Tech - Core Web Vitals", "OnPage - Titles & Meta",
             "OnPage - Headings", "OnPage - Header-Footer H", "OnPage - Images", "Schema - Per Page",
             "Schema - Org & Local", "OffPage - Links & Profiles", "Action Plan", "Data Sources"]
    tab_color = {"Tech": "2F5597", "OnPage": "548235", "Schema": "7030A0", "OffPage": "C55A11",
                 "Issue": "C00000", "Page": "404040", "Action": "BF9000", "Data": "808080"}
    allowed = SCOPES.get(au.scope, {}).get("tabs")
    for name in order:
        if allowed and name not in allowed:
            continue
        headers, rows, widths = T[name]
        if not rows and name not in ("Fix List", "Issue Tracker", "Data Sources"):
            continue  # no problems in this area -> no tab
        ws = wb.create_sheet(name[:31])
        ws.sheet_properties.tabColor = tab_color.get(name.split(" ")[0], "808080")
        ws.append(headers)
        for cell in ws[1]:
            cell.fill, cell.font = hdr_fill, hdr_font
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        if not rows:
            ws.append(["No issues found for this check."] + [""] * (len(headers) - 1))
        for row in rows:
            ws.append(row)
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = wrap
                cell.border = border
                v = cell.value
                if isinstance(v, str) and v.startswith(("http://", "https://")) and "\n" not in v and len(v) < 255:
                    cell.hyperlink = v
                    cell.font = Font(color="0563C1", underline="single")
        ws.freeze_panes = "B2" if name in ("Page Inventory", "Issue Tracker", "Fix List") else "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 32
        last = ws.max_row
        # severity colouring + status dropdown
        if "Severity" in headers:
            col = get_column_letter(headers.index("Severity") + 1)
            for sev, color in SEV_FILL.items():
                ws.conditional_formatting.add(f"{col}2:{col}{last}", CellIsRule(
                    operator="equal", formula=[f'"{sev}"'], fill=PatternFill("solid", fgColor=color),
                    font=Font(color="FFFFFF", bold=True) if sev == "Critical" else None))
        if "Result" in headers:
            col = get_column_letter(headers.index("Result") + 1)
            for word, color in (("Fail", "F8CBAD"), ("Warn", "FFE699"), ("Pass", "C6E0B4"), ("Broken", "F8CBAD")):
                ws.conditional_formatting.add(f"{col}2:{col}{last}", CellIsRule(
                    operator="equal", formula=[f'"{word}"'], fill=PatternFill("solid", fgColor=color)))
        for key, color in (("Over 100 KB", "F8CBAD"), ("Missing alt", "FFE699")):
            if key in headers:
                col = get_column_letter(headers.index(key) + 1)
                ws.conditional_formatting.add(f"{col}2:{col}{last}", CellIsRule(
                    operator="equal", formula=['"Yes"'], fill=PatternFill("solid", fgColor=color)))
        if "Status" in headers and name in ("Issue Tracker", "Action Plan", "Fix List"):
            col = get_column_letter(headers.index("Status") + 1)
            dv = DataValidation(type="list", formula1='"Open,In progress,Done,Won\'t fix"', allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"{col}2:{col}{max(last, 2) + 200}")
    wb.save(path)


def write_csvs(T, au, folder):
    os.makedirs(folder, exist_ok=True)
    for name, (headers, rows, _) in T.items():
        fn = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_") + ".csv"
        with open(os.path.join(folder, fn), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(headers)
            w.writerows(rows)


def summary(au):
    """Condensed findings for Claude to read (keeps token use low)."""
    m = au.c["meta"]
    types = Counter(p["page_type"] for p in au.pages)
    out = {"meta": {k: m.get(k) for k in ("start_url", "mode", "pages_parsed", "urls_status_checked", "images_checked", "notes", "platform")},
           "scope": au.scope or "all",
           "page_types": dict(types),
           "crawler_issues": [{"severity": i["severity"], "area": i["area"], "issue": i["issue"], "count": i["count"],
                               "examples": i["urls"][:3], "evidence": i["evidence"]} for i in au.issues],
           "representative_pages": {t: [p["url"] for p in au.pages if p["page_type"] == t][:3] for t in types},
           "pagespeed": [{k: x.get(k) for k in ("url", "perf_score", "error", "top_opportunities")} for x in au.c.get("pagespeed") or []]}
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("crawl")
    ap.add_argument("--analysis")
    ap.add_argument("--out", default="seo-audit.xlsx")
    ap.add_argument("--csv-dir")
    ap.add_argument("--scope", help="claude-seo skill of the audit (seo-technical, seo-schema, ...); limits issues and tabs")
    ap.add_argument("--summary", action="store_true", help="print condensed findings JSON and exit")
    a = ap.parse_args()
    crawl = json.load(open(a.crawl, encoding="utf-8"))
    analysis = json.load(open(a.analysis, encoding="utf-8")) if a.analysis and os.path.exists(a.analysis) else {}
    au = Audit(crawl, analysis)
    au.scope = a.scope if a.scope in SCOPES else None
    au.run()
    if a.summary:
        print(json.dumps(summary(au), indent=1, ensure_ascii=False))
        return
    T = build_tables(au)
    write_xlsx(au, T, a.out)
    if a.csv_dir:
        allowed = SCOPES.get(au.scope, {}).get("tabs")
        write_csvs({k: v for k, v in T.items() if not allowed or k in allowed}, au, a.csv_dir)
    sev = Counter(i["severity"] for i in au.issues)
    allowed = SCOPES.get(au.scope, {}).get("tabs")
    print(json.dumps({"out": a.out, "scope": au.scope or "all", "platform": au.platform, "issues": len(au.issues),
                      "fixes": len(T["Fix List"][1]), "by_severity": dict(sev),
                      "tabs": ["Summary"] + [k for k in T if (not allowed or k in allowed) and
                                             (T[k][1] or k in ("Fix List", "Issue Tracker", "Data Sources"))]}, indent=1))


if __name__ == "__main__":
    main()
