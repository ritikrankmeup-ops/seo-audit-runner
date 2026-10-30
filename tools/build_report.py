#!/usr/bin/env python3
"""
SEO audit workbook (issues only) from crawl.json.

  python3 build_report.py crawl.json --scope seo-technical --analysis analysis.json --rewrites rewrites.json --out audit.xlsx
  python3 build_report.py crawl.json --summary              # condensed findings for Claude (small)
  python3 build_report.py crawl.json --meta-todo todo.jsonl  # titles/descriptions to (re)write, for Claude

rewrites.json (written by Claude):  {"<url>": {"keyword": "...", "title": "...", "description": "..."}}
analysis.json (written by Claude):  {"audit_label": "...", "executive_summary": "...", "notes": {"<tab name>": "..."}}

Every tab lists ONLY things that need action. A tab with nothing to fix is left out.
"""
import argparse
import csv
import json
import os
import re
from collections import Counter, defaultdict
from urllib.parse import urlparse

TITLE_MAX, DESC_MAX = 60, 160
IMG_KB = 100
THIN = {"Landing / service page": 300, "Blog post / article": 300, "Course page": 300, "Location page": 250,
        "Product page": 150, "Homepage": 300, "Product category / shop": 150, "About page": 200}
NOINDEX_CANDIDATE = re.compile(
    r"(privacy|terms|conditions|cookie|refund|return-polic|shipping-polic|cancellation|disclaimer|legal|gdpr|"
    r"thank-?you|success|cart|checkout|my-account|/account|login|log-in|register|signup|sign-up|wp-login|"
    r"/search|[?&]s=|/tag/|/author/|/page/\d+|/feed|/test|dummy|demo|sample-page|hello-world|staging|"
    r"coming-soon|maintenance|/wp-json|/attachment|replytocom)", re.I)
LOCAL_TYPES = {"localbusiness", "dentist", "medicalclinic", "medicalbusiness", "physician", "store", "restaurant",
               "professionalservice", "homeandconstructionbusiness", "legalservice", "automotivebusiness",
               "healthandbeautybusiness", "hairsalon", "beautysalon", "dayspa", "hospital", "lodgingbusiness",
               "hotel", "realestateagent", "financialservice", "accountingservice", "emergencyservice", "plumber",
               "electrician", "generalcontractor", "lawyer", "attorney", "optician"}
ORG_TYPES = {"organization", "corporation", "medicalorganization", "educationalorganization", "collegeoruniversity",
             "school", "highschool", "ngo", "onlinestore", "onlinebusiness", "newsmediaorganization",
             "sportsorganization", "governmentorganization", "researchorganization", "library"} | LOCAL_TYPES
NEEDED_SCHEMA = {  # page type -> schema types that page should carry ("A|B" = either satisfies)
    "Homepage": ["Organization", "WebSite"],
    "Product page": ["Product", "BreadcrumbList"],
    "Product category / shop": ["CollectionPage|ItemList", "BreadcrumbList"],
    "Course page": ["Course", "BreadcrumbList"],
    "Blog post / article": ["Article|BlogPosting|NewsArticle", "BreadcrumbList"],
    "Blog archive / category": ["CollectionPage|Blog", "BreadcrumbList"],
    "Location page": ["LocalBusiness", "BreadcrumbList"],
    "Contact page": ["ContactPage"],
    "About page": ["AboutPage"],
    "Landing / service page": ["Service", "BreadcrumbList"],
}
TAB_ORDER = ["Redirects", "4xx Errors", "Indexability", "Sitemap Audit", "Core Web Vitals", "Mobile Friendliness",
             "Canonicalisation", "Duplicate URLs", "Missing Titles", "Meta Over Limit", "Duplicate Meta",
             "Missing H1-H2", "Multiple H1", "Thin Content", "Orphan Pages", "Images", "Schema"]
SCOPE_TABS = {
    "seo-technical": ["Redirects", "4xx Errors", "Indexability", "Sitemap Audit", "Core Web Vitals",
                      "Mobile Friendliness", "Canonicalisation", "Duplicate URLs", "Orphan Pages"],
    "seo-sitemap": ["Sitemap Audit", "Indexability", "Redirects", "4xx Errors", "Orphan Pages", "Duplicate URLs"],
    "seo-content": ["Missing Titles", "Meta Over Limit", "Duplicate Meta", "Missing H1-H2", "Multiple H1", "Thin Content"],
    "seo-schema": ["Schema"],
    "seo-images": ["Images", "Core Web Vitals"],
    "seo-local": ["Schema", "Missing Titles", "Meta Over Limit", "Duplicate Meta"],
    "seo-geo": ["Schema", "Missing H1-H2", "Multiple H1", "Thin Content", "Meta Over Limit"],
}
TAB_INFO = {  # what to do, who, priority
    "Redirects": ("Change each anchor (on the listed page) to point at the New URL; remove old URLs from the sitemap.", "Content team", "High"),
    "4xx Errors": ("Restore the page, or 301 it to the closest live page, and fix links to it.", "Developer", "High"),
    "Indexability": ("'Marked noindex': confirm each should stay hidden. 'Should be noindex': add noindex to these pages.", "SEO", "Medium"),
    "Sitemap Audit": ("Remove every coloured (irrelevant) URL from the XML sitemap. See the colour legend below.", "Developer", "High"),
    "Core Web Vitals": ("Improve pages scoring under 90 - mobile first (images, scripts, caching).", "Developer", "High"),
    "Mobile Friendliness": ("Fix every check marked Fail / Warn.", "Developer", "High"),
    "Canonicalisation": ("All URL versions must 301 to one version; each page needs a self-referencing canonical.", "Developer", "High"),
    "Duplicate URLs": ("Keep one URL per page: 301 or canonical the others, and fix internal links.", "Developer", "High"),
    "Missing Titles": ("Add a unique title (50-60 characters) - suggestion included.", "Content team", "High"),
    "Meta Over Limit": ("Replace with the NEW title / description (review first).", "Content team", "Medium"),
    "Duplicate Meta": ("Give every URL its own title and description.", "Content team", "Medium"),
    "Missing H1-H2": ("Add one H1 (main topic) and H2 sub-headings.", "Content team", "High"),
    "Multiple H1": ("Keep one H1; change the others to H2.", "Content team", "Medium"),
    "Thin Content": ("Expand with useful content (FAQs, details, proof).", "Content team", "Medium"),
    "Orphan Pages": ("Add an internal link to each page from the suggested page.", "Content team", "Medium"),
    "Images": ("Compress images over 100 KB (WebP) and add alt text.", "Content team", "Medium"),
    "Schema": ("Add the schema types listed under 'Need to add'.", "Developer", "Medium"),
}
PRIO_FILL = {"High": "F4B084", "Medium": "FFE699", "Low": "C6E0B4"}
SITEMAP_COLORS = {"Redirects": "F8CBAD", "Broken (4xx/5xx)": "FF7C80", "Noindex": "FF7C80",
                  "Canonical points elsewhere": "FFE699", "Should not be indexed": "FFE699",
                  "Listed more than once": "D9D9D9", "No response": "D9D9D9"}


def types_of(node):
    t = node.get("@type")
    return [str(x) for x in (t if isinstance(t, list) else [t]) if x]


def page_types(p):
    s = set()
    for n in p.get("schema", []):
        s.update(types_of(n))
    return s | set(p.get("microdata") or [])


def is_noindex(p):
    return "noindex" in ((p.get("meta_robots") or "") + " " + (p.get("x_robots_tag") or "")).lower()


def norm_key(u):
    """Same page, different spelling: lower-case, no trailing slash, no query, no www, no scheme."""
    pr = urlparse(u)
    host = pr.netloc.lower()
    host = host[4:] if host.startswith("www.") else host
    return host + (pr.path.lower().rstrip("/") or "/")


class Report:
    def __init__(self, crawl, analysis=None, rewrites=None, scope=None):
        self.c = crawl
        self.a = analysis or {}
        self.rw = rewrites or {}
        self.scope = scope if scope in SCOPE_TABS else None
        self.P = crawl["pages"]
        self.pages = {p["url"]: p for p in self.P}
        self.status = {s["url"]: s for s in crawl["status"]}
        self.sitemap = {s["url"]: s for s in crawl.get("sitemap_urls", [])}
        self.inlinks = crawl.get("inlinks", {})
        self.m = crawl["meta"]
        self.T = {}

    def anchors(self, target):
        """[(page, anchor text)] for every internal link pointing at target. A link in the header/footer/menu
        that appears on 3+ pages is collapsed into ONE row ('All pages - footer menu (N pages)')."""
        body, tpl = [], defaultdict(set)
        total = self.c.get("inlink_counts", {}).get(target, 0)
        for src in self.inlinks.get(target, []):
            seen = set()
            for ln in self.pages.get(src, {}).get("links", []):
                if ln["href"] != target:
                    continue
                text = ln["text"] or "(image / icon link)"
                if (text, ln["region"]) in seen:
                    continue
                seen.add((text, ln["region"]))
                if ln["region"] in ("header", "footer", "nav", "aside"):
                    tpl[(ln["region"], text)].add(src)
                else:
                    body.append((src, text))
        out = []
        names = {"header": "header", "footer": "footer", "nav": "main menu", "aside": "sidebar"}
        for (region, text), srcs in tpl.items():
            if len(srcs) >= 3:
                n = max(len(srcs), total) if len(srcs) >= 50 else len(srcs)
                out.append((f"All pages - {names[region]} link ({n} pages, edit it once in the {names[region]} template)", text))
            else:
                out.extend((s, text) for s in sorted(srcs))
        return out + body

    def add(self, name, headers, rows, widths, fills=None):
        self.T[name] = {"headers": headers, "rows": rows, "widths": widths, "fills": fills or []}

    # 1. Redirects
    def redirects(self):
        rows, fills = [], []
        for s in self.status.values():
            if not s.get("chain"):
                continue
            old, final = s["url"], s.get("final_url") or ""
            pairs = self.anchors(old)
            if not pairs and old not in self.sitemap:
                continue  # old URL nobody links to and not in the sitemap: no action needed
            if not pairs:
                pairs = [("(not linked internally - listed in the XML sitemap)", "-")]
            for src, text in pairs:
                rows.append([old, s["chain"][0]["status"], final, s.get("final_status") or "", text, src])
                fills.append("FFE699" if len(s["chain"]) > 1 else None)  # yellow = redirect chain (2+ hops)
        self.add("Redirects", ["Old URL", "Old URL status", "New URL (final)", "New URL status", "Anchor text",
                               "Page where the anchor is used"], rows, [55, 9, 55, 9, 35, 55], fills)

    # 2. 4xx
    def errors_4xx(self):
        rows = sorted([s["url"], s.get("final_status")] for s in self.status.values()
                      if 400 <= (s.get("final_status") or 0) < 500)
        self.add("4xx Errors", ["URL", "Status code"], rows, [80, 12])

    # 4. Indexability
    def indexability(self):
        rows = []
        for p in self.P:
            if is_noindex(p):
                src = "meta robots" if "noindex" in (p.get("meta_robots") or "").lower() else "X-Robots-Tag header"
                rows.append(["Marked noindex", p["url"], p["page_type"], f"noindex ({src})",
                             "Yes" if p["url"] in self.sitemap else "No",
                             "Remove from sitemap (conflict)" if p["url"] in self.sitemap else "Confirm it should stay hidden"])
        for p in self.P:
            if is_noindex(p):
                continue
            m = NOINDEX_CANDIDATE.search(p["url"])
            if m or p["page_type"] in ("Legal / policy page", "Utility page (thank-you / cart / account)"):
                reason = m.group(1).strip("/?&=") if m else p["page_type"]
                rows.append(["Should be noindex", p["url"], p["page_type"], f"Indexable ('{reason}' page)",
                             "Yes" if p["url"] in self.sitemap else "No",
                             "Add noindex and remove from sitemap - no search value"])
        self.add("Indexability", ["Type", "URL", "Page type", "Current state", "In sitemap", "Action"], rows,
                 [18, 60, 24, 30, 10, 45])

    # 5. Sitemap (all sitemap URLs, irrelevant ones coloured)
    def sitemap_audit(self):
        rows, fills = [], []
        for s in self.c.get("sitemap_urls", []):
            u = s["url"]
            st = self.status.get(u, {})
            p = self.pages.get(st.get("final_url") or u) or self.pages.get(u)
            reasons = []
            if st.get("chain"):
                reasons.append("Redirects")
            fs = st.get("final_status") or 0
            if fs >= 400:
                reasons.append("Broken (4xx/5xx)")
            if st.get("unreachable"):
                reasons.append("No response")
            if p and is_noindex(p):
                reasons.append("Noindex")
            if p and p.get("canonical") and p["canonical"].rstrip("/") != p["url"].rstrip("/"):
                reasons.append("Canonical points elsewhere")
            if NOINDEX_CANDIDATE.search(u):
                reasons.append("Should not be indexed")
            if s.get("listed", 1) > 1:
                reasons.append("Listed more than once")
            verdict = "OK" if not reasons else ("Check manually" if reasons == ["No response"] else
                                                ("Duplicate entry - keep one" if reasons == ["Listed more than once"]
                                                 else "Irrelevant - remove"))
            rows.append([u, ", ".join(x.rsplit("/", 1)[-1] for x in s["sitemaps"]), st.get("status") or "",
                         (st.get("final_url") or "") if st.get("chain") else "", fs or "", verdict, ", ".join(reasons) or "-"])
            fills.append(SITEMAP_COLORS[reasons[0]] if reasons else None)
        order = sorted(range(len(rows)), key=lambda i: (rows[i][5] == "OK", rows[i][0]))
        rows, fills = [rows[i] for i in order], [fills[i] for i in order]
        if not any(r[5] != "OK" for r in rows):
            rows, fills = [], []
        self.add("Sitemap Audit", ["URL in sitemap", "Sitemap file", "Status", "Redirects to", "Final status",
                                   "Verdict", "Why"], rows, [60, 25, 8, 45, 8, 18, 40], fills)

    # 6. Core Web Vitals (mobile + desktop)
    def cwv(self):
        rows, fills = [], []
        sec = lambda ms: round(ms / 1000, 2) if ms else ""
        for x in self.c.get("pagespeed") or []:
            if x.get("error"):
                rows.append([x["url"], x.get("device", "").title(), "Not measured", "", "", "", "", "", "", "", "", "",
                             x["error"][:150]])
                fills.append(None)
                continue
            f = x.get("field_url") or {}
            fld = lambda v, cat: f"{v} ({cat.title()})" if v is not None and cat else "No real-user data"
            sc = x.get("perf_score") or 0
            rows.append([x["url"], x.get("device", "mobile").title(), sc, sec(x.get("lab_lcp_ms")),
                         round(x["lab_cls"], 3) if x.get("lab_cls") is not None else "", round(x.get("lab_tbt_ms") or 0),
                         sec(x.get("lab_fcp_ms")), sec(x.get("lab_si_ms")), round(x.get("ttfb_ms") or 0),
                         fld(f.get("lcp_ms"), f.get("lcp_cat")), fld(f.get("inp_ms"), f.get("inp_cat")),
                         fld(f.get("cls"), f.get("cls_cat")), "\n".join((x.get("top_opportunities") or [])[:3])])
            fills.append("C6E0B4" if sc >= 90 else ("FFE699" if sc >= 50 else "F8CBAD"))
        if not rows:
            rows, fills = [["(not measured)", "", "", "", "", "", "", "", "", "", "", "",
                            "PageSpeed did not run - add PAGESPEED_API_KEY to the cloud environment"]], [None]
        self.add("Core Web Vitals", ["URL", "Device", "Performance score (0-100)", "LCP (s)", "CLS", "TBT (ms)",
                                     "FCP (s)", "Speed Index (s)", "Server response (ms)", "Real-user LCP ms (p75)",
                                     "Real-user INP ms (p75)", "Real-user CLS (p75)", "Top fixes"],
                 rows, [45, 9, 12, 8, 7, 9, 8, 10, 11, 20, 20, 18, 60], fills)

    # 18. Mobile friendliness (whole site)
    def mobile(self):
        P, rows = self.P, []
        n = len(P) or 1
        no_vp = [p["url"] for p in P if not p.get("viewport")]
        bad_vp = [p["url"] for p in P if p.get("viewport") and "width=device-width" not in p["viewport"].replace(" ", "")]
        zoom = [p["url"] for p in P if p.get("viewport") and
                re.search(r"user-scalable\s*=\s*(no|0)|maximum-scale\s*=\s*1(\.0)?\b", p["viewport"])]
        rows.append(["Viewport meta tag on every page", "Pass" if not no_vp else "Fail",
                     f"{len(no_vp)} of {n} pages missing", "\n".join(no_vp[:5])])
        rows.append(["Viewport uses width=device-width", "Pass" if not bad_vp else "Fail",
                     f"{len(bad_vp)} pages use a fixed width", "\n".join(bad_vp[:5])])
        rows.append(["Users can zoom (no user-scalable=no / maximum-scale=1)", "Pass" if not zoom else "Warn",
                     f"{len(zoom)} pages block pinch-zoom (accessibility)", "\n".join(zoom[:5])])
        mob = [x for x in self.c.get("pagespeed") or [] if x.get("device") == "mobile" and not x.get("error")]
        if mob:
            avg = round(sum(x["perf_score"] for x in mob) / len(mob))
            rows.append(["Mobile speed score (PageSpeed)", "Pass" if avg >= 90 else ("Warn" if avg >= 50 else "Fail"),
                         f"Average {avg}/100 on {len(mob)} page(s)", "\n".join(f"{x['url']}: {x['perf_score']}" for x in mob)])
            fo = (mob[0].get("field_origin") or {}).get("overall")
            if fo:
                rows.append(["Real-user mobile Core Web Vitals (whole site)", "Pass" if fo == "FAST" else "Fail",
                             f"Google CrUX verdict: {fo.title()}", ""])
            for key, label in (("font-size", "Text readable on mobile"), ("tap-targets", "Buttons/links big enough to tap"),
                               ("content-width", "No horizontal scrolling (content fits screen)"),
                               ("uses-responsive-images", "Images sized for mobile")):
                hits = [x for x in mob if key in (x.get("mobile_audits") or {})]
                if hits:
                    bad = [x for x in hits if (x["mobile_audits"][key].get("score") or 0) < 0.9]
                    rows.append([label, "Pass" if not bad else "Fail",
                                 ("; ".join(f"{x['url']}: {x['mobile_audits'][key].get('display') or 'fails'}" for x in bad)[:300]
                                  or f"OK on {len(hits)} page(s)"), ""])
        else:
            rows.append(["Mobile speed score (PageSpeed)", "Not measured", "PageSpeed did not run", ""])
        imgs = [im for p in P for im in p.get("images", []) if im["region"] in ("main", "body")]
        if imgs:
            pct = round(100 * sum(1 for im in imgs if im.get("srcset")) / len(imgs))
            rows.append(["Images use srcset (smaller files on phones)", "Pass" if pct >= 70 else "Warn",
                         f"{pct}% of content images", ""])
        bad = sum(1 for r in rows if r[1] in ("Fail", "Warn"))
        rows.insert(0, ["OVERALL MOBILE-FRIENDLINESS", "Pass" if bad == 0 else ("Needs work" if bad <= 2 else "Fail"),
                        f"{bad} check(s) need attention", ""])
        self.add("Mobile Friendliness", ["Check", "Result", "Detail", "Example pages"], rows, [50, 14, 70, 60])

    # 11. Canonicalisation (site-wide versions + page level)
    def canonical(self):
        rows, fills = [], []
        start = self.m["start_url"]
        pref_host = urlparse(start).netloc
        for v in self.c.get("site_checks", {}).get("url_versions", []):
            target = v.get("canonical_target") or start
            fs, hops, fu = v.get("final_status"), v.get("hops", 0), v.get("final_url") or ""
            if v["url"].rstrip("/") == start.rstrip("/"):
                continue  # this "version" IS the audited URL
            if v.get("error") or (fs and fs >= 400):
                continue  # version does not exist / not reachable = no duplicate
            if hops and urlparse(fu).netloc == pref_host and \
                    (v["version"] == "Uppercase path" or fu.rstrip("/") == target.rstrip("/")):
                continue  # correctly 301s to the preferred version
            res = f"Redirects to {fu} (status {fs}) - not the preferred version" if hops else \
                "Loads separately with status 200 - DUPLICATE VERSION of the site"
            rows.append(["Whole website", v["version"], v["url"], res, f"301 redirect to {target}"])
            fills.append("F8CBAD")
        hashes = defaultdict(list)
        for p in self.P:
            if p.get("content_hash"):
                hashes[p["content_hash"]].append(p["url"])
        for p in self.P:
            c = p.get("canonical")
            dup = [u for u in hashes.get(p.get("content_hash"), []) if u != p["url"]] if p.get("content_hash") else []
            issue = None
            if not c:
                issue = "No canonical tag"
            elif c.rstrip("/") != p["url"].rstrip("/"):
                st = self.status.get(c, {})
                if st and ((st.get("final_status") or 0) >= 300 or st.get("chain")):
                    issue = f"Canonical points to a redirecting/broken URL: {c}"
                elif not dup:
                    issue = f"Canonical points to a different page: {c}"
            if dup and (not c or c.rstrip("/") == p["url"].rstrip("/")):
                issue = (issue + "; " if issue else "") + "Same content as other URL(s) - each says it is the original"
            if issue:
                rows.append(["Page", p["page_type"], p["url"],
                             issue + (f"\nOther versions: {', '.join(dup[:3])}" if dup else ""),
                             "Choose ONE version; point every version's canonical (or a 301) to it" if dup else
                             "Add a self-referencing canonical" if not c else "Set the canonical to the page's own URL"])
                fills.append(None)
        self.add("Canonicalisation", ["Level", "Version / page type", "URL", "Problem", "Fix"], rows,
                 [16, 32, 60, 70, 45], fills)

    # 16. Duplicate URLs
    def duplicate_urls(self):
        live = {u for u, s in self.status.items() if s.get("final_status") == 200 and not s.get("chain")} | set(self.pages)
        by_key = defaultdict(set)
        for u in live:
            by_key[norm_key(u)].add(u)
        groups = [sorted(v) for v in by_key.values() if len(v) > 1]
        seen = {frozenset(g) for g in groups}
        by_hash = defaultdict(set)
        for p in self.P:
            if p.get("content_hash"):
                by_hash[p["content_hash"]].add(p["url"])
        groups += [sorted(v) for v in by_hash.values() if len(v) > 1 and frozenset(v) not in seen]
        rows = []
        for g, grp in enumerate(groups, 1):
            why = "Same address, different spelling (case / slash / parameter / www)" \
                if len({norm_key(x) for x in grp}) == 1 else "Identical page content on different URLs"
            for u in grp:
                rows.append([g, u, 200, self.pages.get(u, {}).get("canonical") or "(none)", why])
        self.add("Duplicate URLs", ["Group", "URL", "Status", "Canonical tag on page", "Why duplicate"], rows,
                 [7, 70, 8, 60, 50])

    # 15. Missing titles
    def missing_titles(self):
        rows = []
        for p in self.P:
            if not (p.get("title") or "").strip():
                h1 = next((h["text"] for h in p["headings"] if h["level"] == 1 and h["text"]), "")
                rows.append([p["url"], p["page_type"], h1, self.rw.get(p["url"], {}).get("title", "")])
        self.add("Missing Titles", ["URL", "Page type", "H1 on page", "Suggested title"], rows, [65, 24, 45, 60])

    # 7. Meta over limit (+ AI rewrites)
    def meta_todo(self):
        out = []
        for p in self.P:
            t, d = (p.get("title") or "").strip(), (p.get("meta_description") or "").strip()
            if len(t) > TITLE_MAX or len(d) > DESC_MAX:
                out.append({"url": p["url"], "page_type": p["page_type"], "title": t, "title_len": len(t),
                            "description": d, "desc_len": len(d),
                            "h1": next((h["text"] for h in p["headings"] if h["level"] == 1 and h["text"]), ""),
                            "snippet": (p.get("snippet") or "")[:300]})
        prio = {"Homepage": 0, "Landing / service page": 1, "Course page": 1, "Product page": 1, "Location page": 2,
                "Product category / shop": 2, "About page": 3, "Contact page": 3, "Blog post / article": 4}
        out.sort(key=lambda r: (prio.get(r["page_type"], 5), r["url"]))
        return out

    def meta_over(self):
        rows = []
        for r in self.meta_todo():
            rw = self.rw.get(r["url"], {})
            nt, nd = rw.get("title", ""), rw.get("description", "")
            t_new = nt or ("Pending (not rewritten in this run)" if r["title_len"] > TITLE_MAX else "OK - keep")
            d_new = nd or ("Pending (not rewritten in this run)" if r["desc_len"] > DESC_MAX else "OK - keep")
            rows.append([r["url"], r["page_type"], rw.get("keyword", ""), r["title"], r["title_len"], t_new,
                         len(nt) if nt else "", r["description"], r["desc_len"], d_new, len(nd) if nd else ""])
        self.add("Meta Over Limit", ["URL", "Page type", "Primary keyword", "Current title", "Title length",
                                     "NEW SEO title", "New length", "Current description", "Desc length",
                                     "NEW meta description", "New length"],
                 rows, [50, 20, 22, 45, 8, 50, 8, 55, 8, 60, 8])

    # 12. Duplicate meta
    def duplicate_meta(self):
        rows = []
        for kind, key in (("Title", "title"), ("Meta description", "meta_description")):
            groups = defaultdict(list)
            for p in self.P:
                v = (p.get(key) or "").strip()
                if v:
                    groups[v.lower()].append((v, p["url"]))
            for _, lst in sorted(groups.items(), key=lambda kv: -len(kv[1])):
                if len(lst) > 1:
                    rows.append([kind, lst[0][0], len(lst), "\n".join(u for _, u in lst[:30]) +
                                 (f"\n... +{len(lst) - 30} more" if len(lst) > 30 else "")])
        self.add("Duplicate Meta", ["Type", "Duplicated text", "Pages using it", "URLs"], rows, [16, 70, 10, 75])

    # 9. Missing H1 / H2
    def missing_h(self):
        rows = []
        no_h2_ok = ("Utility page (thank-you / cart / account)", "Legal / policy page", "Contact page")
        for p in self.P:
            hs = p.get("headings", [])
            h1 = [h for h in hs if h["level"] == 1 and h["text"]]
            h2 = [h for h in hs if h["level"] == 2 and h["text"] and h["region"] in ("main", "body")]
            miss = (["H1"] if not h1 else []) + (["H2"] if not h2 and p["page_type"] not in no_h2_ok else [])
            if not miss:
                continue
            first = next((h for h in hs if h["region"] in ("main", "body") and h["text"]), None)
            sug = []
            if "H1" in miss:
                base = first["text"] if first else re.split(r"\s[|\-–]\s", p.get("title") or "")[0]
                sug.append(f"Add H1: '{base[:70]}'" + (f" (now an H{first['level']})" if first else ""))
            if "H2" in miss:
                sug.append("Break the content into sections with H2 sub-headings (e.g. benefits, process, FAQs)")
            rows.append([p["url"], p["page_type"], " + ".join(miss), h1[0]["text"] if h1 else "(none)", "; ".join(sug)])
        self.add("Missing H1-H2", ["URL", "Page type", "Missing", "Current H1", "What to add"], rows, [60, 22, 10, 40, 70])

    # 10. Multiple H1
    def multi_h1(self):
        rows = []
        for p in self.P:
            h1 = [h for h in p.get("headings", []) if h["level"] == 1]
            if len(h1) > 1:
                rows.append([p["url"], len(h1), "\n".join(f"{h['text'] or '(empty)'}  [{h['region']}]" for h in h1),
                             f"Keep '{(h1[0]['text'] or '')[:50]}' as H1; change the other {len(h1) - 1} to H2"])
        self.add("Multiple H1", ["URL", "H1 count", "H1 texts [where on page]", "Fix"], rows, [60, 9, 60, 55])

    # 13. Thin content
    def thin(self):
        rows = [[p["url"], p["page_type"], p["word_count"], THIN[p["page_type"]], THIN[p["page_type"]] - p["word_count"]]
                for p in self.P if p["page_type"] in THIN and p["word_count"] < THIN[p["page_type"]] and not is_noindex(p)]
        rows.sort(key=lambda r: r[2])
        self.add("Thin Content", ["URL", "Page type", "Words (main content)", "Minimum for this type", "Words to add"],
                 rows, [70, 24, 12, 12, 10])

    # 14. Orphans
    def orphans(self):
        rows = []
        live = {urlparse(u).path: u for u in self.pages}
        for u in self.sitemap:
            if u not in self.pages or self.inlinks.get(u) or u.rstrip("/") == self.m["start_url"].rstrip("/") \
                    or is_noindex(self.pages[u]) or self.pages[u]["page_type"] in (
                        "Legal / policy page", "Utility page (thank-you / cart / account)"):
                continue
            parts, hub = urlparse(u).path.rstrip("/").split("/"), None
            while len(parts) > 1 and not hub:
                parts = parts[:-1]
                hub = live.get("/".join(parts) + "/") or live.get("/".join(parts) or "/")
            rows.append([u, self.pages[u]["page_type"], "Yes", 0, hub or "a related category / hub page"])
        self.add("Orphan Pages", ["URL", "Page type", "In sitemap", "Internal links to it", "Add a link from"],
                 rows, [70, 24, 10, 10, 60])

    # 8. Images
    def images(self):
        alt_missing = defaultdict(set)
        for p in self.P:
            for im in p.get("images", []):
                if im["region"] in ("main", "body") and (not im.get("has_alt_attr") or not (im.get("alt") or "").strip()):
                    alt_missing[im["src"]].add(p["url"])
        sizes = {i["url"]: i for i in self.c.get("images", [])}
        rows = []
        for u in set(sizes) | set(alt_missing):
            i = sizes.get(u, {})
            kb = round(i["bytes"] / 1024, 1) if i.get("bytes") else None
            big, noalt = bool(kb and kb > IMG_KB), u in alt_missing
            if not (big or noalt):
                continue
            pages = sorted(alt_missing[u]) if noalt else (i.get("pages") or [])
            rows.append([u, kb if kb is not None else "", "Yes" if big else "No", "Yes" if noalt else "No",
                         len(pages) or i.get("page_count", 0), "\n".join(pages[:5])])
        rows.sort(key=lambda r: -(r[1] or 0))
        self.add("Images", ["Image URL", "Size (KB)", "Over 100 KB", "Missing alt text", "Pages using it",
                            "Pages (examples)"], rows, [70, 10, 10, 12, 10, 60])

    # 17. Schema (present vs needed only)
    def schema(self):
        rows, P = [], self.P
        all_types = Counter(t for p in P for t in page_types(p))
        has = lambda names: any(t.lower() in names for t in all_types)
        local = any(p.get("tel_links") for p in P) or any(p["page_type"] == "Location page" for p in P)
        site = [("Organization (or subtype)", has(ORG_TYPES), True, "Brand entity: name, logo, social profiles"),
                ("WebSite", has({"website"}), True, "Site name shown in Google results"),
                ("LocalBusiness (or subtype)", has(LOCAL_TYPES), local, "Address, phone, hours for local searches"),
                ("BreadcrumbList", has({"breadcrumblist"}), len(P) > 1, "Breadcrumb trail in Google results")]
        for name, present, needed, why in site:
            if needed or present:
                rows.append(["WHOLE WEBSITE", name, "Present" if present else "Missing", "-" if present else "Add", why])
        rows.append(["WHOLE WEBSITE", "All schema types found", ", ".join(f"{t} ({n} pages)" for t, n in all_types.most_common(12))
                     or "None", "", ""])
        for p in P:
            have = page_types(p)
            hl = {t.lower() for t in have}
            need = []
            for exp in NEEDED_SCHEMA.get(p["page_type"], []):
                opts = {o.lower() for o in exp.split("|")} | (LOCAL_TYPES if exp == "LocalBusiness" else set())
                if not hl & opts:
                    need.append(exp.split("|")[0])
            if p.get("faq_like") and "faqpage" not in hl:
                need.append("FAQPage")
            if need:
                rows.append([p["url"], p["page_type"], ", ".join(sorted(have)) or "None", ", ".join(need),
                             "FAQ section found on page" if "FAQPage" in need else ""])
        if not any(r[3] == "Add" or (r[0] != "WHOLE WEBSITE") for r in rows):
            rows = []
        self.add("Schema", ["Page", "Page type / schema", "Schema present", "Need to add", "Why"], rows,
                 [55, 28, 50, 35, 40])

    def build(self):
        for f in (self.redirects, self.errors_4xx, self.indexability, self.sitemap_audit, self.cwv, self.mobile,
                  self.canonical, self.duplicate_urls, self.missing_titles, self.meta_over, self.duplicate_meta,
                  self.missing_h, self.multi_h1, self.thin, self.orphans, self.images, self.schema):
            f()
        allowed = SCOPE_TABS.get(self.scope, TAB_ORDER)
        self.tabs = [t for t in TAB_ORDER if t in allowed and self.T[t]["rows"]]
        return self


def write_xlsx(R, path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    m = R.m
    wb = Workbook()
    hdr_fill, hdr_font = PatternFill("solid", fgColor="1F3864"), Font(bold=True, color="FFFFFF")
    wrap = Alignment(wrap_text=True, vertical="top")
    border = Border(bottom=Side(style="thin", color="D9D9D9"))

    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = f"SEO Audit: {m['domain']}"
    ws["A1"].font = Font(bold=True, size=16, color="1F3864")
    plat = (m.get("platform") or {}).get("label", "")
    ws["A2"] = f"{R.a.get('audit_label', '')} | {m['start_url']} | Platform: {plat} | {m['crawled_at']}"
    ws["A2"].font = Font(italic=True, color="595959")
    r = 4
    if R.a.get("executive_summary"):
        ws.cell(r, 1, "Summary").font = Font(bold=True, size=12)
        ws.cell(r + 1, 1, R.a["executive_summary"]).alignment = wrap
        ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=5)
        ws.row_dimensions[r + 1].height = 95
        r += 3
    ws.cell(r, 1, f"Crawled {m['pages_parsed']} pages | {len(R.sitemap)} sitemap URLs | "
                  f"{m.get('urls_status_checked', 0)} URLs checked"
                  + (f" | NOTE: {'; '.join(m.get('notes', []))}" if m.get("notes") else "")).alignment = wrap
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.row_dimensions[r].height = 32
    r += 2
    for j, h in enumerate(["Tab", "Items to fix", "Priority", "What to do", "Who"], 1):
        c = ws.cell(r, j, h)
        c.fill, c.font = hdr_fill, hdr_font
    r += 1
    for t in R.tabs:
        rows = R.T[t]["rows"]
        n = sum(1 for x in rows if x[5] != "OK") if t == "Sitemap Audit" else \
            sum(1 for x in rows[1:] if x[1] in ("Fail", "Warn")) if t == "Mobile Friendliness" else len(rows)
        what, who, pr = TAB_INFO[t]
        note = (R.a.get("notes") or {}).get(t)
        ws.cell(r, 1, t).hyperlink = f"#'{t}'!A1"
        ws.cell(r, 1).font = Font(color="0563C1", underline="single")
        ws.cell(r, 2, n)
        ws.cell(r, 3, pr).fill = PatternFill("solid", fgColor=PRIO_FILL[pr])
        ws.cell(r, 4, what + (f"\n{note}" if note else "")).alignment = wrap
        ws.cell(r, 5, who)
        r += 1
    if not R.tabs:
        ws.cell(r, 1, "No issues found in this audit scope.")
    if "Sitemap Audit" in R.tabs:
        r += 1
        ws.cell(r, 1, "Sitemap colour legend").font = Font(bold=True)
        r += 1
        for k, col in list(SITEMAP_COLORS.items()) + [("OK (no colour)", "FFFFFF")]:
            ws.cell(r, 1, k).fill = PatternFill("solid", fgColor=col)
            r += 1
    for col, w in zip("ABCDE", (28, 13, 11, 85, 16)):
        ws.column_dimensions[col].width = w

    for name in R.tabs:
        t = R.T[name]
        ws = wb.create_sheet(name[:31])
        ws.append(t["headers"])
        for c in ws[1]:
            c.fill, c.font = hdr_fill, hdr_font
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for i, row in enumerate(t["rows"]):
            ws.append(row)
            fill = t["fills"][i] if i < len(t["fills"]) else None
            for c in ws[ws.max_row]:
                c.alignment, c.border = wrap, border
                if fill:
                    c.fill = PatternFill("solid", fgColor=fill)
                v = c.value
                if isinstance(v, str) and v.startswith(("http://", "https://")) and "\n" not in v and len(v) < 255:
                    c.hyperlink = v
                    c.font = Font(color="0563C1", underline="single")
        for i, w in enumerate(t["widths"], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 34
        if "Result" in t["headers"]:
            col = t["headers"].index("Result") + 1
            for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
                for c in row:
                    color = {"Pass": "C6E0B4", "Fail": "F8CBAD", "Warn": "FFE699", "Needs work": "FFE699"}.get(c.value)
                    if color:
                        c.fill = PatternFill("solid", fgColor=color)
    wb.save(path)


def write_csvs(R, folder):
    os.makedirs(folder, exist_ok=True)
    for name in R.tabs:
        t = R.T[name]
        with open(os.path.join(folder, re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_") + ".csv"), "w",
                  newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(t["headers"])
            w.writerows(t["rows"])


def summary(R):
    m = R.m
    return {"meta": {k: m.get(k) for k in ("start_url", "mode", "pages_parsed", "urls_status_checked", "notes", "platform")},
            "scope": R.scope or "all",
            "page_types": dict(Counter(p["page_type"] for p in R.P)),
            "tabs": {t: len(R.T[t]["rows"]) for t in R.tabs},
            "examples": {t: [r[:4] for r in R.T[t]["rows"][:3]] for t in R.tabs},
            "meta_to_rewrite": len(R.meta_todo()) if {"Meta Over Limit", "Missing Titles"} & set(R.tabs) else 0,
            "pagespeed": [{k: x.get(k) for k in ("url", "device", "perf_score", "error")} for x in R.c.get("pagespeed") or []]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("crawl")
    ap.add_argument("--analysis")
    ap.add_argument("--rewrites")
    ap.add_argument("--scope")
    ap.add_argument("--out", default="seo-audit.xlsx")
    ap.add_argument("--csv-dir")
    ap.add_argument("--summary", action="store_true")
    ap.add_argument("--meta-todo", help="write missing titles + titles/descriptions over the limit to this JSONL")
    a = ap.parse_args()
    def load(p):
        try:
            return json.load(open(p, encoding="utf-8")) if p and os.path.exists(p) else {}
        except (json.JSONDecodeError, OSError):
            print(f"WARNING: could not read {p} - continuing without it")
            return {}
    R = Report(json.load(open(a.crawl, encoding="utf-8")), load(a.analysis), load(a.rewrites), a.scope).build()
    if a.summary:
        print(json.dumps(summary(R), indent=1, ensure_ascii=False))
        return
    if a.meta_todo:
        missing = [{"url": p["url"], "page_type": p["page_type"], "title": "", "title_len": 0,
                    "description": p.get("meta_description") or "", "desc_len": len(p.get("meta_description") or ""),
                    "h1": next((h["text"] for h in p["headings"] if h["level"] == 1 and h["text"]), ""),
                    "snippet": (p.get("snippet") or "")[:300]} for p in R.P if not (p.get("title") or "").strip()]
        todo = missing + R.meta_todo()
        with open(a.meta_todo, "w", encoding="utf-8") as f:
            for row in todo:
                f.write(json.dumps(row, ensure_ascii=False) + "\n")
        print(json.dumps({"meta_todo": a.meta_todo, "rows": len(todo)}))
        return
    write_xlsx(R, a.out)
    if a.csv_dir:
        write_csvs(R, a.csv_dir)
    print(json.dumps({"out": a.out, "scope": R.scope or "all", "tabs": {t: len(R.T[t]["rows"]) for t in R.tabs}}, indent=1))


if __name__ == "__main__":
    main()
