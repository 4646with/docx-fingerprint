#!/usr/bin/env python3
"""docx-fingerprint: inspect and rewrite OOXML metadata to eliminate cross-file fingerprints."""

import argparse
import hashlib
import os
import random
import re
import sys
import zipfile
from datetime import datetime


def random_rsid():
    return "%08X" % random.randint(0x10000000, 0xFFFFFFFF)


def read_meta(path):
    """Read metadata from a .docx file."""
    z = zipfile.ZipFile(path)
    meta = {"file": os.path.basename(path), "size_kb": os.path.getsize(path) // 1024}

    for xml_name, fields in [
        ("docProps/core.xml", [
            "creator", "lastModifiedBy", "created", "modified",
            "revision", "title", "description", "subject", "keywords",
        ]),
        ("docProps/app.xml", [
            "Application", "AppVersion", "Template", "Company",
            "Manager", "TotalTime",
        ]),
    ]:
        try:
            xml = z.read(xml_name).decode("utf-8")
        except KeyError:
            continue
        for tag in fields:
            patterns = [
                rf"<[^>]*{tag}[^>]*>([^<]+)<",
            ]
            for pat in patterns:
                m = re.search(pat, xml)
                if m and m.group(1).strip():
                    meta[tag] = m.group(1).strip()
                    break

    # RSIDs
    rsids = set()
    for internal in ["word/settings.xml", "word/document.xml"]:
        try:
            xml = z.read(internal).decode("utf-8")
            rsids.update(re.findall(r'w:rsid[^=]*="([^"]+)"', xml))
        except KeyError:
            pass
    meta["rsid_count"] = len(rsids)
    meta["rsid_sample"] = sorted(rsids)[:3]
    meta["rsid_all"] = rsids

    # Special files
    for name in z.namelist():
        if any(k in name.lower() for k in ["comment", "people", "custom.xml"]):
            meta.setdefault("extra_files", []).append(name)

    # Embedded authors in tracked changes
    try:
        doc_xml = z.read("word/document.xml").decode("utf-8")
        authors = set(re.findall(r'w:author="([^"]+)"', doc_xml))
        if authors:
            meta["embedded_authors"] = sorted(authors)
    except KeyError:
        pass

    z.close()
    return meta


def print_meta(meta, verbose=False):
    """Pretty-print metadata."""
    print(f"\n  {meta['file']}  ({meta['size_kb']} KB)")
    print(f"  {'─' * 40}")
    for key in [
        "creator", "lastModifiedBy", "created", "modified", "revision",
        "title", "description", "Application", "AppVersion", "Template",
        "Company", "Manager", "TotalTime",
    ]:
        if key in meta:
            print(f"  {key:20s}: {meta[key]}")
    print(f"  {'RSID count':20s}: {meta['rsid_count']}")
    print(f"  {'RSID sample':20s}: {meta['rsid_sample']}")
    if meta.get("extra_files"):
        print(f"  {'extra files':20s}: {meta['extra_files']}")
    if meta.get("embedded_authors"):
        print(f"  {'embedded authors':20s}: {meta['embedded_authors']}")


def cross_check(metas):
    """Check for overlapping fingerprints between files."""
    names = [m["file"] for m in metas]
    issues = []

    for i in range(len(metas)):
        for j in range(i + 1, len(metas)):
            a, b = metas[i], metas[j]
            na, nb = names[i], names[j]

            if a.get("creator") and a.get("creator") == b.get("creator"):
                issues.append(f"creator identical: {na} ↔ {nb} = {a['creator']}")

            if a.get("lastModifiedBy") and a.get("lastModifiedBy") == b.get("lastModifiedBy"):
                issues.append(f"lastModifiedBy identical: {na} ↔ {nb} = {a['lastModifiedBy']}")

            if a.get("Application") and a.get("Application") == b.get("Application"):
                issues.append(f"Application identical: {na} ↔ {nb} = {a['Application']}")

            rsid_overlap = a.get("rsid_all", set()) & b.get("rsid_all", set())
            if rsid_overlap:
                issues.append(f"RSID overlap: {na} ↔ {nb} = {len(rsid_overlap)} shared ({sorted(rsid_overlap)[:3]}...)")

    return issues


def patch_file(path, output=None, **overrides):
    """
    Rewrite metadata fields in a .docx file.

    Supported overrides:
        author, last_modified_by, application, app_version,
        template, total_time, title, description, company,
        created, modified, revision, randomize_rsids (bool),
        strip_comments (bool), strip_custom (bool)
    """
    output = output or path
    tmp = output + ".tmp"

    new_rsids_settings = [random_rsid() for _ in range(30)]
    new_rsids_doc = [random_rsid() for _ in range(50)]

    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            # Skip files if requested
            if overrides.get("strip_comments") and "comment" in item.filename.lower():
                continue
            if overrides.get("strip_comments") and "people" in item.filename.lower():
                continue
            if overrides.get("strip_custom") and item.filename == "docProps/custom.xml":
                continue

            data = zin.read(item.filename)

            if item.filename == "docProps/core.xml":
                text = data.decode("utf-8")
                replacements = {
                    "creator": overrides.get("author"),
                    "lastModifiedBy": overrides.get("last_modified_by"),
                    "title": overrides.get("title"),
                    "description": overrides.get("description"),
                    "revision": overrides.get("revision"),
                }
                for tag, val in replacements.items():
                    if val is not None:
                        # Handle both dc: and cp: namespaces
                        for ns in ["dc", "cp", "dcterms"]:
                            pat = rf"(<{ns}:{tag}[^>]*>)[^<]*(</)"
                            text = re.sub(pat, rf"\g<1>{val}\2", text)

                for tag, val in [
                    ("created", overrides.get("created")),
                    ("modified", overrides.get("modified")),
                ]:
                    if val is not None:
                        text = re.sub(
                            rf"(<dcterms:{tag}[^>]*>)[^<]*(</)",
                            rf"\g<1>{val}\2",
                            text,
                        )

                # Remove python-docx traces
                text = text.replace("generated by python-docx", "")
                data = text.encode("utf-8")

            elif item.filename == "docProps/app.xml":
                text = data.decode("utf-8")
                app_replacements = {
                    "Application": overrides.get("application"),
                    "AppVersion": overrides.get("app_version"),
                    "Template": overrides.get("template"),
                    "TotalTime": overrides.get("total_time"),
                    "Company": overrides.get("company"),
                }
                for tag, val in app_replacements.items():
                    if val is not None:
                        if val == "":
                            text = re.sub(rf"<{tag}>[^<]*</{tag}>", "", text)
                        else:
                            if f"<{tag}>" in text:
                                text = re.sub(rf"<{tag}>[^<]*</{tag}>", f"<{tag}>{val}</{tag}>", text)
                            else:
                                text = text.replace("</Properties>", f"<{tag}>{val}</{tag}></Properties>")
                data = text.encode("utf-8")

            elif overrides.get("randomize_rsids"):
                if item.filename == "word/settings.xml":
                    text = data.decode("utf-8")
                    idx = [0]
                    def _repl_s(m, _idx=idx, _rsids=new_rsids_settings):
                        i = _idx[0] % len(_rsids)
                        _idx[0] += 1
                        return 'w:rsid%s="%s"' % (m.group(1), _rsids[i])
                    text = re.sub(r'w:rsid([^=]*)="([^"]+)"', _repl_s, text)
                    data = text.encode("utf-8")

                elif item.filename == "word/document.xml":
                    text = data.decode("utf-8")
                    idx2 = [0]
                    def _repl_d(m, _idx=idx2, _rsids=new_rsids_doc):
                        i = _idx[0] % len(_rsids)
                        _idx[0] += 1
                        return 'w:rsid%s="%s"' % (m.group(1), _rsids[i])
                    text = re.sub(r'w:rsid([^=]*)="([^"]+)"', _repl_d, text)
                    data = text.encode("utf-8")

            zout.writestr(item, data)

    os.replace(tmp, output)
    return output


def cmd_inspect(args):
    metas = [read_meta(f) for f in args.files]
    for m in metas:
        print_meta(m, verbose=args.verbose)

    if len(metas) > 1:
        print("\n  Cross-file check")
        print(f"  {'─' * 40}")
        issues = cross_check(metas)
        if issues:
            for issue in issues:
                print(f"  ⚠  {issue}")
        else:
            print("  ✓  No overlapping fingerprints found")
    print()


def cmd_patch(args):
    overrides = {}
    if args.author:
        overrides["author"] = args.author
        overrides["last_modified_by"] = args.author
    if args.last_modified_by:
        overrides["last_modified_by"] = args.last_modified_by
    if args.application:
        overrides["application"] = args.application
    if args.app_version:
        overrides["app_version"] = args.app_version
    if args.template is not None:
        overrides["template"] = args.template
    if args.total_time is not None:
        overrides["total_time"] = args.total_time
    if args.title:
        overrides["title"] = args.title
    if args.company is not None:
        overrides["company"] = args.company
    if args.created:
        overrides["created"] = args.created
    if args.modified:
        overrides["modified"] = args.modified
    if args.revision:
        overrides["revision"] = args.revision
    if args.description is not None:
        overrides["description"] = args.description
    if args.randomize_rsids:
        overrides["randomize_rsids"] = True
    if args.strip_comments:
        overrides["strip_comments"] = True
    if args.strip_custom:
        overrides["strip_custom"] = True

    if not overrides:
        print("No changes specified. Use --help to see options.", file=sys.stderr)
        sys.exit(1)

    for f in args.files:
        out = args.output if args.output and len(args.files) == 1 else f
        patch_file(f, out, **overrides)
        print(f"Patched: {out}")

    # Show result
    metas = [read_meta(args.output if args.output and len(args.files) == 1 else f) for f in args.files]
    for m in metas:
        print_meta(m)
    print()


def main():
    parser = argparse.ArgumentParser(
        prog="docx-fingerprint",
        description="Inspect and rewrite OOXML (.docx) metadata to eliminate cross-file fingerprints.",
    )
    sub = parser.add_subparsers(dest="command")

    # inspect
    p_inspect = sub.add_parser("inspect", help="Show metadata of one or more .docx files and check for overlaps")
    p_inspect.add_argument("files", nargs="+", help=".docx files to inspect")
    p_inspect.add_argument("-v", "--verbose", action="store_true")

    # patch
    p_patch = sub.add_parser("patch", help="Rewrite metadata fields in a .docx file")
    p_patch.add_argument("files", nargs="+", help=".docx files to patch")
    p_patch.add_argument("-o", "--output", help="Output path (only for single file)")
    p_patch.add_argument("--author", help="Set creator and lastModifiedBy")
    p_patch.add_argument("--last-modified-by", help="Set lastModifiedBy only (overrides --author)")
    p_patch.add_argument("--application", help='e.g. "Microsoft Office Word", "WPS Office", "LibreOffice Writer"')
    p_patch.add_argument("--app-version", help='e.g. "16.0000", "7.6000"')
    p_patch.add_argument("--template", help='e.g. "Normal.dotm", or "" to remove')
    p_patch.add_argument("--total-time", help="Editing time in minutes")
    p_patch.add_argument("--title", help="Document title")
    p_patch.add_argument("--description", help='Document description, or "" to clear')
    p_patch.add_argument("--company", help='Company field, or "" to clear')
    p_patch.add_argument("--created", help="Created timestamp, e.g. 2026-11-18T14:22:00Z")
    p_patch.add_argument("--modified", help="Modified timestamp")
    p_patch.add_argument("--revision", help="Revision number")
    p_patch.add_argument("--randomize-rsids", action="store_true", help="Replace all RSIDs with random values")
    p_patch.add_argument("--strip-comments", action="store_true", help="Remove comments and people.xml")
    p_patch.add_argument("--strip-custom", action="store_true", help="Remove custom.xml properties")

    args = parser.parse_args()
    if args.command == "inspect":
        cmd_inspect(args)
    elif args.command == "patch":
        cmd_patch(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
