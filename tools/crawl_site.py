#!/usr/bin/env python3
"""
SEO site crawler -> crawl.json (deterministic data collection for the audit workbook).

Collects, per URL:
  - status code, full redirect chain (which URL redirects to which), final URL, response time
  - title / meta description / canonical / meta robots / X-Robots-Tag / lang / viewport / OG
  - every heading in document order with its region (header / nav / main / aside / footer)
  - images (src, alt, region) and their file sizes (HEAD / streamed GET)
  - internal + external links (anchor, rel, region) and their status
  - JSON-LD schema blocks (parsed, @graph flattened, invalid JSON recorded)
  - page type (home / landing-service / product / product-category / blog-post / archive / ...)
Plus site-level: robots.txt, sitemaps (index-aware), http->https and www checks.

Usage:
  python3 crawl_site.py https://example.com --max-pages 200 --out crawl.json
  python3 crawl_site.py https://example.com/page/ --single   # one page + its links/images
"""
import argparse
import os
import ipaddress
import json
import re
import socket
import sys
import threading
import time
import urllib.robotparser as robotparser
import xml.etree.ElementTree as ET
from collections import defaultdict, deque
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from urllib.parse import urljoin, urlparse, urlunparse, urldefrag

import requests
from bs4 import BeautifulSoup

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0 Safari/537.36 SEOAuditBot/1.0")
HTML_TYPES = ("text/html", "application/xhtml+xml")
REGION_RE = re.compile(
    r"(^|[\s_-])(site-header|masthead|header|top-bar|topbar|navbar|nav|menu|"
    r"site-footer|footer|colophon|sidebar|widget-area|elementor-location-header|"
    r"elementor-location-footer)($|[\s_-])", re.I)
SKIP_EXT = re.compile(r"\.(pdf|zip|rar|7z|gz|mp4|mp3|avi|mov|webm|docx?|xlsx?|pptx?|exe|dmg|apk|"
                      r"jpg|jpeg|png|gif|webp|avif|svg|ico|css|js|woff2?|ttf|eot|xml|txt|json)$", re.I)
PHONE_RE = re.compile(r"(?:\+?\d[\d\s\-().]{7,}\d)")


# ----------------------------------------------------------------------------- helpers
def now_iso():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")


def norm_url(url, base=None):
    """Absolute URL, no fragment, lowercase scheme/host, default ports removed."""
    if not url:
        return None
    url = url.strip()
    if url.startswith(("mailto:", "tel:", "javascript:", "data:", "#", "sms:", "whatsapp:")):
        return None
    if base:
        url = urljoin(base, url)
    url, _ = urldefrag(url)
    p = urlparse(url)
    if p.scheme not in ("http", "https") or not p.netloc:
        return None
    netloc = p.netloc.lower()
    if (p.scheme == "http" and netloc.endswith(":80")) or (p.scheme == "https" and netloc.endswith(":443")):
        netloc = netloc.rsplit(":", 1)[0]
    path = p.path or "/"
    return urlunparse((p.scheme.lower(), netloc, path, "", p.query, ""))


def bare_host(host):
    host = (host or "").lower().split(":")[0]
    return host[4:] if host.startswith("www.") else host


def is_private_target(url):
    """True if the URL points at a private/loopback/link-local address (SSRF guard)."""
    host = urlparse(url).hostname or ""
    if host in ("localhost",) or host.endswith(".local") or host.endswith(".internal"):
        return True
    try:
        ip = ipaddress.ip_address(host)
        return ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved
    except ValueError:
        pass
    try:
        infos = socket.getaddrinfo(host, None)
    except Exception:
        return False  # cannot resolve locally (proxy resolves) -> let the proxy decide
    for info in infos:
        try:
            ip = ipaddress.ip_address(info[4][0])
        except ValueError:
            continue
        if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved:
            return True
    return False


def region_of(tag):
    """Which template region a tag sits in: header / nav / footer / aside / main / body."""
    for parent in tag.parents:
        name = getattr(parent, "name", None)
        if not name or name in ("html", "[document]"):
            break
        if name in ("header", "footer", "nav", "aside"):
            return name
        role = (parent.get("role") or "").lower()
        if role in ("banner",):
            return "header"
        if role in ("contentinfo",):
            return "footer"
        if role in ("navigation",):
            return "nav"
        tokens = " ".join(parent.get("class") or []) + " " + (parent.get("id") or "")
        m = REGION_RE.search(tokens)
        if m:
            word = m.group(2).lower()
            if "footer" in word or word == "colophon":
                return "footer"
            if word in ("nav", "navbar", "menu"):
                return "nav"
            if word in ("sidebar", "widget-area"):
                return "aside"
            return "header"
        if name == "main" or role == "main":
            return "main"
    return "body"


# ----------------------------------------------------------------------------- crawler
class Crawler:
    def __init__(self, start, max_pages, single, allow_private, delay, workers,
                 max_minutes, check_external, max_images, respect_robots, timeout):
        self.start = norm_url(start)
        if not self.start:
            raise SystemExit(f"Invalid start URL: {start}")
        self.root_host = bare_host(urlparse(self.start).netloc)
        self.max_pages = 1 if single else max_pages
        self.single = single
        self.allow_private = allow_private
        self.delay = delay
        self.workers = workers
        self.deadline = time.time() + max_minutes * 60
        self.check_external = check_external
        self.max_images = max_images
        self.respect_robots = respect_robots
        self.timeout = timeout
        self.local = threading.local()
        self.lock = threading.Lock()
        self.status = {}          # url -> status record (pages, links, sitemap urls)
        self.pages = {}           # url -> parsed page data (final 200 HTML pages)
        self.inlinks = defaultdict(set)   # target url -> set(source pages)
        self.found_in = defaultdict(set)  # url -> {"start","sitemap","internal link"}
        self.images = {}          # img url -> record
        self.image_pages = defaultdict(set)
        self.external = {}        # ext url -> record
        self.ext_sources = defaultdict(set)
        self.sitemap_urls = []
        self.sitemaps = []
        self.robots = None
        self.robots_txt = ""
        self.site_checks = {}
        self.notes = []

    # -- http
    def session(self):
        s = getattr(self.local, "s", None)
        if s is None:
            s = requests.Session()
            s.headers.update({"User-Agent": UA, "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
                              "Accept-Language": "en-US,en;q=0.9"})
            self.local.s = s
        return s

    def is_internal(self, url):
        return bare_host(urlparse(url).netloc) == self.root_host

    def time_left(self):
        return time.time() < self.deadline

    def fetch(self, url, method="GET", want_body=True, max_hops=10):
        """Follow redirects manually so every hop is recorded."""
        chain, cur, t0 = [], url, time.time()
        rec = {"url": url, "status": None, "final_url": None, "final_status": None, "chain": [],
               "content_type": "", "elapsed_ms": None, "error": None, "headers": {}, "body": None,
               "bytes": None}
        for _ in range(max_hops + 1):
            if not self.allow_private and is_private_target(cur):
                rec["error"] = "blocked: private/local address"
                break
            try:
                time.sleep(self.delay)
                r = self.session().request(method, cur, allow_redirects=False, timeout=self.timeout,
                                           stream=not want_body)
            except requests.exceptions.SSLError as e:
                rec["error"] = f"SSL error: {str(e)[:160]}"
                break
            except requests.exceptions.Timeout:
                rec["error"] = f"timeout after {self.timeout}s"
                break
            except requests.exceptions.RequestException as e:
                rec["error"] = f"connection error: {str(e)[:160]}"
                break
            if rec["status"] is None:
                rec["status"] = r.status_code
            if r.status_code in (301, 302, 303, 307, 308) and r.headers.get("Location"):
                nxt = norm_url(r.headers["Location"], cur)
                chain.append({"url": cur, "status": r.status_code, "to": nxt})
                r.close()
                if not nxt or nxt in [c["url"] for c in chain]:
                    rec["error"] = "redirect loop" if nxt else "invalid redirect target"
                    break
                cur = nxt
                continue
            rec["final_url"] = cur
            rec["final_status"] = r.status_code
            rec["content_type"] = (r.headers.get("Content-Type") or "").split(";")[0].strip().lower()
            rec["headers"] = {k.lower(): v for k, v in r.headers.items()
                              if k.lower() in ("x-robots-tag", "cache-control", "content-length",
                                               "server", "strict-transport-security", "link")}
            cl = r.headers.get("Content-Length")
            rec["bytes"] = int(cl) if cl and cl.isdigit() else None
            if want_body and rec["content_type"] in HTML_TYPES and r.status_code == 200:
                r.encoding = r.encoding or r.apparent_encoding
                rec["body"] = r.text
                rec["bytes"] = rec["bytes"] or len(r.content)
            r.close()
            break
        else:
            rec["error"] = f"more than {max_hops} redirects"
        rec["chain"] = chain
        rec["elapsed_ms"] = int((time.time() - t0) * 1000)
        if rec["final_status"] is None and rec["status"] is not None and not chain:
            rec["final_status"] = rec["status"]
        return rec

    def size_of(self, url):
        """Image/resource size in bytes: HEAD Content-Length, else streamed GET."""
        if not self.allow_private and is_private_target(url):
            return {"status": None, "bytes": None, "type": "", "error": "blocked: private address"}
        try:
            r = self.session().head(url, allow_redirects=True, timeout=self.timeout)
            cl = r.headers.get("Content-Length")
            ctype = (r.headers.get("Content-Type") or "").split(";")[0].lower()
            if r.status_code < 400 and cl and cl.isdigit() and int(cl) > 0:
                return {"status": r.status_code, "bytes": int(cl), "type": ctype, "error": None}
            r = self.session().get(url, stream=True, allow_redirects=True, timeout=self.timeout)
            total = 0
            for chunk in r.iter_content(65536):
                total += len(chunk)
                if total > 20 * 1024 * 1024:
                    break
            ctype = (r.headers.get("Content-Type") or ctype).split(";")[0].lower()
            st = r.status_code
            r.close()
            return {"status": st, "bytes": total if st < 400 else None, "type": ctype, "error": None}
        except requests.exceptions.RequestException as e:
            return {"status": None, "bytes": None, "type": "", "error": str(e)[:160]}

    # -- site level
    def load_robots(self):
        p = urlparse(self.start)
        robots_url = f"{p.scheme}://{p.netloc}/robots.txt"
        rec = self.fetch(robots_url, want_body=False)
        txt = ""
        try:
            r = self.session().get(robots_url, timeout=self.timeout)
            if r.status_code == 200:
                txt = r.text[:20000]
        except requests.exceptions.RequestException:
            pass
        self.robots_txt = txt
        self.site_checks["robots_txt"] = {"url": robots_url, "status": rec["final_status"] or rec["status"],
                                          "found": bool(txt), "length": len(txt)}
        if txt:
            rp = robotparser.RobotFileParser()
            rp.parse(txt.splitlines())
            self.robots = rp
        return [ln.split(":", 1)[1].strip() for ln in txt.splitlines()
                if ln.lower().startswith("sitemap:")]

    def allowed(self, url):
        if not self.respect_robots or not self.robots:
            return True
        try:
            return self.robots.can_fetch("*", url)
        except Exception:
            return True

    def load_sitemaps(self, declared):
        p = urlparse(self.start)
        base = f"{p.scheme}://{p.netloc}"
        queue = deque(declared or [])
        for guess in ("/sitemap_index.xml", "/sitemap.xml", "/wp-sitemap.xml"):
            queue.append(base + guess)
        seen, urls = set(), []
        while queue and len(seen) < 60 and len(urls) < 20000 and self.time_left():
            sm = norm_url(queue.popleft())
            if not sm or sm in seen:
                continue
            seen.add(sm)
            try:
                r = self.session().get(sm, timeout=self.timeout)
            except requests.exceptions.RequestException as e:
                self.sitemaps.append({"url": sm, "status": None, "type": "error", "count": 0, "error": str(e)[:120]})
                continue
            if r.status_code != 200 or b"<" not in r.content[:500]:
                if sm in [norm_url(d) for d in declared]:
                    self.sitemaps.append({"url": sm, "status": r.status_code, "type": "missing", "count": 0, "error": ""})
                continue
            try:
                root = ET.fromstring(r.content)
            except ET.ParseError as e:
                self.sitemaps.append({"url": sm, "status": 200, "type": "invalid XML", "count": 0, "error": str(e)[:120]})
                continue
            tag = root.tag.split("}")[-1]
            locs = [el.text.strip() for el in root.iter() if el.tag.split("}")[-1] == "loc" and el.text]
            if tag == "sitemapindex":
                self.sitemaps.append({"url": sm, "status": 200, "type": "index", "count": len(locs), "error": ""})
                queue.extend(locs)
            else:
                self.sitemaps.append({"url": sm, "status": 200, "type": "urlset", "count": len(locs), "error": ""})
                for loc in locs:
                    u = norm_url(loc)
                    if u:
                        urls.append((u, sm))
        dup = defaultdict(list)
        for u, sm in urls:
            dup[u].append(sm)
        self.sitemap_urls = [{"url": u, "sitemaps": sorted(set(s)), "listed": len(s)} for u, s in dup.items()]
        for u in dup:
            self.found_in[u].add("sitemap")

    def site_level_checks(self):
        p = urlparse(self.start)
        host = p.netloc
        alt_host = host[4:] if host.startswith("www.") else "www." + host
        checks = {}
        for label, url in (("http_root", f"http://{host}/"), ("alt_www_root", f"https://{alt_host}/"),
                           ("http_alt_www_root", f"http://{alt_host}/")):
            rec = self.fetch(url, want_body=False)
            checks[label] = {"url": url, "status": rec["status"], "final_url": rec["final_url"],
                             "final_status": rec["final_status"], "hops": len(rec["chain"]),
                             "error": rec["error"]}
        nf = self.fetch(f"{p.scheme}://{host}/this-page-should-not-exist-seo-audit-404-check/", want_body=False)
        checks["soft_404_test"] = {"status": nf["final_status"] or nf["status"], "error": nf["error"]}
        self.site_checks.update(checks)

    # -- page parsing
    def parse_page(self, url, rec):
        soup = BeautifulSoup(rec["body"], "lxml")
        head = soup.head or soup
        def meta(name=None, prop=None):
            tag = head.find("meta", attrs={"name": re.compile(f"^{name}$", re.I)}) if name else \
                head.find("meta", attrs={"property": re.compile(f"^{prop}$", re.I)})
            return (tag.get("content") or "").strip() if tag else None

        title_tag = soup.find("title")
        title = title_tag.get_text(" ", strip=True) if title_tag else None
        canon = head.find("link", rel=lambda v: v and "canonical" in (v if isinstance(v, list) else [v]))
        canonical = norm_url(canon.get("href"), url) if canon and canon.get("href") else None
        hreflang = [{"lang": l.get("hreflang"), "href": norm_url(l.get("href"), url)}
                    for l in head.find_all("link", hreflang=True)]
        body = soup.body or soup
        body_classes = " ".join(body.get("class") or []) if soup.body else ""

        # headings (document order) with region
        headings = []
        for h in soup.find_all(re.compile(r"^h[1-6]$")):
            text = h.get_text(" ", strip=True)
            headings.append({"level": int(h.name[1]), "text": text[:200], "region": region_of(h),
                             "empty": not text})

        # images
        imgs = []
        for img in soup.find_all("img"):
            src = img.get("data-src") or img.get("data-lazy-src") or img.get("data-original") or img.get("src")
            if (not src or src.startswith("data:")) and img.get("srcset"):
                src = img.get("srcset").split(",")[0].strip().split(" ")[0]
            if not src or src.startswith("data:"):
                continue
            src = norm_url(src, url)
            if not src:
                continue
            imgs.append({"src": src, "alt": img.get("alt"), "has_alt_attr": img.has_attr("alt"),
                         "width": img.get("width"), "height": img.get("height"),
                         "loading": img.get("loading"), "region": region_of(img)})

        # links
        links = []
        for a in soup.find_all("a", href=True):
            href = norm_url(a["href"], url)
            if not href:
                continue
            rel = " ".join(a.get("rel") or []).lower()
            links.append({"href": href, "text": a.get_text(" ", strip=True)[:120], "rel": rel,
                          "internal": self.is_internal(href), "region": region_of(a)})

        # schema
        schema_nodes, schema_errors = [], []
        for s in soup.find_all("script", type=re.compile("ld\\+json", re.I)):
            raw = (s.string or s.get_text() or "").strip()
            if not raw:
                continue
            try:
                data = json.loads(raw)
            except json.JSONDecodeError as e:
                schema_errors.append(f"Invalid JSON-LD: {str(e)[:100]}")
                continue
            stack = data if isinstance(data, list) else [data]
            while stack:
                node = stack.pop(0)
                if isinstance(node, dict):
                    if "@graph" in node and isinstance(node["@graph"], list):
                        stack.extend(node["@graph"])
                    if "@type" in node:
                        schema_nodes.append(node)
                elif isinstance(node, list):
                    stack.extend(node)
        microdata = sorted({(m.get("itemtype") or "").split("/")[-1] for m in soup.find_all(attrs={"itemtype": True})})

        # text + misc
        for t in soup(["script", "style", "noscript", "template", "svg"]):
            t.decompose()
        main_text_parts = []
        for el in (soup.body or soup).find_all(string=True):
            if region_of(el.parent) in ("main", "body"):
                main_text_parts.append(el)
        main_text = " ".join(" ".join(main_text_parts).split())
        tel_links = sorted({a["href"][4:].strip() for a in soup.find_all("a", href=True)
                            if a["href"].lower().startswith("tel:")})
        mixed = []
        if url.startswith("https://"):
            for tag, attr in (("img", "src"), ("script", "src"), ("link", "href"), ("iframe", "src"),
                              ("source", "src"), ("video", "src")):
                for el in BeautifulSoup(rec["body"], "lxml").find_all(tag):
                    v = el.get(attr) or ""
                    if v.startswith("http://"):
                        mixed.append(v[:200])
        html_lang = soup.html.get("lang") if soup.html else None
        page = {
            "url": url, "status": 200, "elapsed_ms": rec["elapsed_ms"], "bytes": rec["bytes"],
            "title": title, "meta_description": meta(name="description"),
            "meta_robots": meta(name="robots"), "x_robots_tag": rec["headers"].get("x-robots-tag"),
            "canonical": canonical, "lang": html_lang, "viewport": meta(name="viewport"),
            "og_title": meta(prop="og:title"), "og_description": meta(prop="og:description"),
            "og_image": meta(prop="og:image"), "twitter_card": meta(name="twitter:card"),
            "hreflang": hreflang, "headings": headings, "images": imgs, "links": links,
            "schema": schema_nodes, "schema_errors": schema_errors, "microdata": [m for m in microdata if m],
            "word_count": len(main_text.split()), "body_classes": body_classes,
            "tel_links": tel_links, "mixed_content": mixed[:20],
            "has_form": bool(soup.find("form")),
            "add_to_cart": bool(re.search(r"add[ _-]?to[ _-]?cart|single_add_to_cart|buy now",
                                          rec["body"][:400000], re.I)),
            "cache_control": rec["headers"].get("cache-control"),
        }
        page["page_type"] = classify(page)
        return page

    # -- main loop
    def crawl(self):
        declared = self.load_robots()
        if not self.single:
            self.load_sitemaps(declared)
            self.site_level_checks()
        self.found_in[self.start].add("start URL")
        queue = deque([self.start])
        queued = {self.start}
        if not self.single:
            for item in self.sitemap_urls:
                if item["url"] not in queued and self.is_internal(item["url"]):
                    queue.append(item["url"])
                    queued.add(item["url"])
        fetched_html = 0
        with ThreadPoolExecutor(self.workers) as pool:
            while queue and fetched_html < self.max_pages and self.time_left():
                batch = []
                while queue and len(batch) < self.workers and fetched_html + len(batch) < self.max_pages:
                    u = queue.popleft()
                    if SKIP_EXT.search(urlparse(u).path):
                        continue
                    if not self.allowed(u):
                        self.status[u] = {"url": u, "status": None, "final_status": None, "chain": [],
                                          "final_url": None, "error": "blocked by robots.txt",
                                          "elapsed_ms": None, "content_type": ""}
                        continue
                    batch.append(u)
                for u, rec in zip(batch, pool.map(self.fetch, batch)):
                    body = rec.pop("body", None)
                    self.status[u] = rec
                    final = rec["final_url"]
                    if rec["chain"] and final and final not in self.status:
                        self.found_in[final].add("redirect target")
                    if body and rec["final_status"] == 200 and final and self.is_internal(final) \
                            and final not in self.pages:
                        rec["body"] = body
                        page = self.parse_page(final, rec)
                        rec.pop("body", None)
                        self.pages[final] = page
                        fetched_html += 1
                        for ln in page["links"]:
                            if ln["internal"]:
                                self.inlinks[ln["href"]].add(final)
                                self.found_in[ln["href"]].add("internal link")
                                if not self.single and ln["href"] not in queued:
                                    queued.add(ln["href"])
                                    queue.append(ln["href"])
                            else:
                                self.ext_sources[ln["href"]].add(final)
                        for im in page["images"]:
                            self.image_pages[im["src"]].add(final)
        if not self.time_left():
            self.notes.append("Crawl stopped early: time budget reached.")
        if queue and fetched_html >= self.max_pages:
            self.notes.append(f"Page limit reached ({self.max_pages} HTML pages). "
                              f"{len(queue)} more URLs were discovered but not crawled.")
        self.status_check_remaining()
        self.check_images()
        self.check_externals()

    def status_check_remaining(self):
        """Status-check internal link targets + sitemap URLs not crawled (no parsing)."""
        targets = [u for u in list(self.inlinks) + [s["url"] for s in self.sitemap_urls]
                   if u not in self.status and not SKIP_EXT.search(urlparse(u).path)]
        targets = list(dict.fromkeys(targets))[:4000]
        with ThreadPoolExecutor(self.workers) as pool:
            for u, rec in zip(targets, pool.map(lambda x: self.fetch(x, want_body=False), targets)):
                if not self.time_left():
                    break
                rec.pop("body", None)
                self.status[u] = rec

    def check_images(self):
        imgs = list(self.image_pages)[: self.max_images]
        with ThreadPoolExecutor(self.workers) as pool:
            for u, res in zip(imgs, pool.map(self.size_of, imgs)):
                self.images[u] = res
        if len(self.image_pages) > self.max_images:
            self.notes.append(f"Image size checked for {self.max_images} of {len(self.image_pages)} images.")

    def check_externals(self):
        if not self.check_external:
            return
        exts = list(self.ext_sources)[: self.check_external]
        def chk(u):
            try:
                r = self.session().head(u, allow_redirects=True, timeout=self.timeout)
                st = r.status_code
                if st in (405, 403, 400):
                    r = self.session().get(u, allow_redirects=True, timeout=self.timeout, stream=True)
                    st = r.status_code
                    r.close()
                return {"status": st, "final_url": r.url, "error": None}
            except requests.exceptions.RequestException as e:
                return {"status": None, "final_url": None, "error": str(e)[:120]}
        with ThreadPoolExecutor(self.workers) as pool:
            for u, res in zip(exts, pool.map(chk, exts)):
                self.external[u] = res

    def pagespeed(self, key, limit=6):
        """Lab + field Core Web Vitals (mobile) for the start URL + one page per page type."""
        picks, seen = [self.start], {"Homepage"}
        for p in self.pages.values():
            if p["page_type"] not in seen and len(picks) < limit:
                seen.add(p["page_type"]); picks.append(p["url"])
        picks = list(dict.fromkeys(picks))[:limit]
        def run(u):
            api = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
            try:
                r = requests.get(api, params={"url": u, "strategy": "mobile", "category": "performance", "key": key},
                                 timeout=120)
                if r.status_code != 200:
                    return {"url": u, "error": f"PSI HTTP {r.status_code}: {r.text[:120]}"}
                d = r.json()
                lh = d.get("lighthouseResult", {})
                au = lh.get("audits", {})
                num = lambda k: au.get(k, {}).get("numericValue")
                def field(src):
                    m = (d.get(src) or {}).get("metrics") or {}
                    g = lambda k: (m.get(k) or {})
                    return {"lcp_ms": g("LARGEST_CONTENTFUL_PAINT_MS").get("percentile"),
                            "lcp_cat": g("LARGEST_CONTENTFUL_PAINT_MS").get("category"),
                            "inp_ms": g("INTERACTION_TO_NEXT_PAINT").get("percentile"),
                            "inp_cat": g("INTERACTION_TO_NEXT_PAINT").get("category"),
                            "cls": (g("CUMULATIVE_LAYOUT_SHIFT_SCORE").get("percentile") or 0) / 100
                            if g("CUMULATIVE_LAYOUT_SHIFT_SCORE") else None,
                            "cls_cat": g("CUMULATIVE_LAYOUT_SHIFT_SCORE").get("category"),
                            "overall": (d.get(src) or {}).get("overall_category")}
                opps = sorted(((a.get("details") or {}).get("overallSavingsMs") or 0, a.get("title"))
                              for a in au.values() if (a.get("details") or {}).get("type") == "opportunity")
                return {"url": u, "error": None,
                        "perf_score": round((lh.get("categories", {}).get("performance", {}).get("score") or 0) * 100),
                        "lab_lcp_ms": num("largest-contentful-paint"), "lab_cls": num("cumulative-layout-shift"),
                        "lab_tbt_ms": num("total-blocking-time"), "lab_fcp_ms": num("first-contentful-paint"),
                        "field_url": field("loadingExperience"), "field_origin": field("originLoadingExperience"),
                        "top_opportunities": [f"{t} (~{int(ms)} ms)" for ms, t in reversed(opps[-4:]) if ms]}
            except Exception as e:  # noqa: BLE001
                return {"url": u, "error": str(e)[:160]}
        with ThreadPoolExecutor(3) as pool:
            return list(pool.map(run, picks))

    def result(self):
        return {
            "meta": {"start_url": self.start, "domain": self.root_host, "crawled_at": now_iso(),
                     "mode": "single page" if self.single else "site crawl",
                     "max_pages": self.max_pages, "pages_parsed": len(self.pages),
                     "urls_status_checked": len(self.status), "images_found": len(self.image_pages),
                     "images_checked": len(self.images), "external_links": len(self.ext_sources),
                     "notes": self.notes, "user_agent": UA},
            "site_checks": self.site_checks,
            "robots_txt": self.robots_txt,
            "pagespeed": getattr(self, "psi", []),
            "sitemaps": self.sitemaps,
            "sitemap_urls": self.sitemap_urls,
            "pages": list(self.pages.values()),
            "status": list(self.status.values()),
            "inlinks": {k: sorted(v)[:50] for k, v in self.inlinks.items()},
            "inlink_counts": {k: len(v) for k, v in self.inlinks.items()},
            "found_in": {k: sorted(v) for k, v in self.found_in.items()},
            "images": [{"url": u, **self.images.get(u, {}), "pages": sorted(self.image_pages[u])[:20],
                        "page_count": len(self.image_pages[u])} for u in self.image_pages],
            "external_links": [{"url": u, **self.external.get(u, {}), "pages": sorted(self.ext_sources[u])[:20],
                                "page_count": len(self.ext_sources[u])} for u in self.ext_sources],
        }


# ----------------------------------------------------------------------------- page type
def schema_types(page):
    types = set()
    for n in page.get("schema", []):
        t = n.get("@type")
        for x in (t if isinstance(t, list) else [t]):
            if x:
                types.add(str(x))
    return types | set(page.get("microdata") or [])


def classify(page):
    url = page["url"]
    path = urlparse(url).path.lower().rstrip("/") or "/"
    bc = page.get("body_classes", "").lower()
    types = {t.lower() for t in schema_types(page)}
    if path == "/" or "home" in bc.split():
        return "Homepage"
    if "product" in types or "single-product" in bc or re.search(r"/(product|products|p|item)/[^/]+$", path):
        return "Product page"
    if re.search(r"/(product-category|collections?|shop|category/products?)(/|$)", path) or \
            "tax-product_cat" in bc or "woocommerce-shop" in bc:
        return "Product category / shop"
    if {"blogposting", "article", "newsarticle"} & types or "single-post" in bc or \
            re.search(r"/(blog|news|articles?|insights?|resources)/.+", path):
        return "Blog post / article"
    if re.search(r"/(category|tag|author|page/\d+|archive)(/|$)", path) or "archive" in bc.split() \
            or "blog" in bc.split() or path in ("/blog", "/news"):
        return "Blog archive / category"
    if re.search(r"(contact|get-in-touch)", path):
        return "Contact page"
    if re.search(r"(about|team|our-story|who-we-are)", path):
        return "About page"
    if re.search(r"(privacy|terms|cookie|refund|disclaimer|policy|shipping)", path):
        return "Legal / policy page"
    if re.search(r"(location|near-me|city|clinic|branch|store-locator|areas?-we-serve)", path) or \
            {"localbusiness", "dentist", "store", "restaurant", "medicalclinic"} & types:
        return "Location page"
    if re.search(r"(thank-you|thankyou|success|cart|checkout|my-account|login|wp-admin)", path):
        return "Utility page (thank-you / cart / account)"
    return "Landing / service page"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("url")
    ap.add_argument("--out", default="crawl.json")
    ap.add_argument("--max-pages", type=int, default=200)
    ap.add_argument("--single", action="store_true", help="audit one page only")
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--delay", type=float, default=0.2, help="seconds between requests per worker")
    ap.add_argument("--timeout", type=int, default=20)
    ap.add_argument("--max-minutes", type=float, default=15)
    ap.add_argument("--max-images", type=int, default=800)
    ap.add_argument("--check-external", type=int, default=300, help="max external links to status-check (0=off)")
    ap.add_argument("--ignore-robots", action="store_true")
    ap.add_argument("--no-psi", action="store_true", help="skip PageSpeed Insights")
    ap.add_argument("--allow-private", action="store_true", help="testing only: allow localhost targets")
    a = ap.parse_args()
    c = Crawler(a.url, a.max_pages, a.single, a.allow_private, a.delay, a.workers, a.max_minutes,
                a.check_external, a.max_images, not a.ignore_robots, a.timeout)
    t0 = time.time()
    c.crawl()
    key = os.environ.get("PAGESPEED_API_KEY", "").strip()
    if key and not a.no_psi:
        c.psi = c.pagespeed(key)
    elif not a.no_psi:
        c.notes.append("Core Web Vitals not measured: set PAGESPEED_API_KEY (free Google API key).")
    res = c.result()
    res["meta"]["duration_sec"] = int(time.time() - t0)
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False)
    m = res["meta"]
    print(json.dumps({"out": a.out, "pages_parsed": m["pages_parsed"], "urls_checked": m["urls_status_checked"],
                      "images_checked": m["images_checked"], "duration_sec": m["duration_sec"],
                      "notes": m["notes"]}, indent=1))


if __name__ == "__main__":
    main()
