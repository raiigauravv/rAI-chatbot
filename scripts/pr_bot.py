#!/usr/bin/env python3
import json
import os
import re
import sys
import base64
from typing import List, Dict, Tuple

import requests
import nbformat

API = "https://api.github.com"
SESSION = requests.Session()
SESSION.headers.update({
    "Accept": "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
})

MARKER = "<!-- pr-comment-bot:ds -->"
DATA_EXTS = {".csv", ".parquet", ".json", ".xlsx", ".feather", ".pkl"}
IPYNB_EXT = ".ipynb"

def github_env() -> Tuple[str, str, str, str]:
    token = os.getenv("GITHUB_TOKEN")
    repo = os.getenv("GITHUB_REPOSITORY")  # owner/repo
    event_path = os.getenv("GITHUB_EVENT_PATH")
    if not all([token, repo, event_path]):
        print("Missing required environment variables.", file=sys.stderr)
        sys.exit(1)
    with open(event_path, "r", encoding="utf-8") as f:
        event = json.load(f)
    pr = event.get("pull_request", {})
    pr_number = pr.get("number") or event.get("number")
    head_sha = (pr.get("head") or {}).get("sha")
    if not pr_number or not head_sha:
        print("Not a pull_request event or missing head SHA.", file=sys.stderr)
        sys.exit(0)
    SESSION.headers["Authorization"] = f"Bearer {token}"
    return repo, str(pr_number), head_sha, (pr.get("base") or {}).get("repo", {}).get("full_name") or repo

def list_pr_files(repo: str, pr_number: str) -> List[Dict]:
    files = []
    page = 1
    while True:
        url = f"{API}/repos/{repo}/pulls/{pr_number}/files"
        r = SESSION.get(url, params={"per_page": 100, "page": page})
        r.raise_for_status()
        chunk = r.json()
        files.extend(chunk)
        if len(chunk) < 100:
            break
        page += 1
    return files

def fetch_file_content(repo: str, path: str, ref: str) -> bytes:
    # GitHub contents API: returns base64 for files
    url = f"{API}/repos/{repo}/contents/{path}"
    r = SESSION.get(url, params={"ref": ref})
    if r.status_code == 404:
        return b""
    r.raise_for_status()
    data = r.json()
    if data.get("type") != "file":
        return b""
    # guard on huge files
    size = data.get("size", 0)
    # Skip > 1.5MB notebooks to avoid timeouts; comment will say it was skipped
    if size and size > 1_500_000:
        return b"__SKIPPED_TOO_LARGE__"
    content = data.get("content", "")
    encoding = data.get("encoding", "base64")
    if encoding == "base64":
        return base64.b64decode(content.encode("utf-8"))
    # fallback
    return content.encode("utf-8")

def notebook_has_outputs(nb_bytes: bytes) -> Tuple[bool, str]:
    if nb_bytes == b"__SKIPPED_TOO_LARGE__":
        return (False, "skipped (file too large for inline check)")
    try:
        nb = nbformat.reads(nb_bytes.decode("utf-8"), as_version=4)
    except Exception as e:
        return (False, f"unable to parse ({e.__class__.__name__})")
    try:
        for cell in nb.cells:
            if cell.get("cell_type") == "code":
                outs = cell.get("outputs") or []
                if len(outs) > 0:
                    return (True, "outputs present")
                # Optional: executed but no outputs
                if cell.get("execution_count"):
                    # Some teams consider any execution_count as “dirty”
                    pass
        return (False, "outputs cleared ✅")
    except Exception as e:
        return (False, f"check error ({e.__class__.__name__})")

def build_comment(repo: str, pr_number: str, head_sha: str,
                  ipynb_results: List[Tuple[str, str]], data_files: List[str],
                  metrics: Dict) -> str:
    lines = []
    lines.append(MARKER)
    lines.append(f"### NoteGuardian 🛡️")
    lines.append(f"_Analyzed PR #{pr_number} @ `{head_sha[:7]}`_")
    lines.append("")
    if ipynb_results:
        lines.append("#### Notebooks changed")
        lines.append("| File | Status |")
        lines.append("|------|--------|")
        for path, status in ipynb_results:
            lines.append(f"| `{path}` | {status} |")
        lines.append("")
        lines.append("> Tip: Clear outputs via `jupyter nbconvert --ClearOutputPreprocessor.enabled=True --inplace your_notebook.ipynb`")
        lines.append("> Or add a pre-commit hook: `nbstripout`")
        lines.append("")
    if data_files:
        lines.append("#### Data files changed")
        for p in data_files:
            lines.append(f"- `{p}`")
        lines.append("")

    if metrics:
        lines.append("#### Model metrics")
        # render a simple table if numeric
        numeric = {k: v for k, v in metrics.items() if isinstance(v, (int, float))}
        if numeric:
            lines.append("| Metric | Value |")
            lines.append("|--------|-------|")
            for k, v in numeric.items():
                lines.append(f"| {k} | {v:.4f} |")
        else:
            # fallback to json code
            lines.append("```json")
            lines.append(json.dumps(metrics, indent=2))
            lines.append("```")
        lines.append("")

    if not (ipynb_results or data_files or metrics):
        lines.append("_No notebooks, data files, or metrics detected in this PR._")
    return "\n".join(lines)

def find_existing_comment(repo: str, issue_number: str) -> str:
    # Return comment_id to update if our marker is found; else ""
    page = 1
    while True:
        url = f"{API}/repos/{repo}/issues/{issue_number}/comments"
        r = SESSION.get(url, params={"per_page": 100, "page": page})
        r.raise_for_status()
        comments = r.json()
        if not comments:
            break
        for c in reversed(comments):
            if isinstance(c.get("body"), str) and MARKER in c["body"]:
                return str(c["id"])
        if len(comments) < 100:
            break
        page += 1
    return ""

def post_or_update_comment(repo: str, issue_number: str, body: str, comment_id: str = ""):
    if comment_id:
        url = f"{API}/repos/{repo}/issues/comments/{comment_id}"
        r = SESSION.patch(url, json={"body": body})
        r.raise_for_status()
        print(f"Updated comment {comment_id}")
    else:
        url = f"{API}/repos/{repo}/issues/{issue_number}/comments"
        r = SESSION.post(url, json={"body": body})
        r.raise_for_status()
        print(f"Posted new comment {r.json().get('id')}")

def load_metrics_if_any() -> Dict:
    # If your CI generated metrics.json in the workspace, include it.
    try:
        if os.path.exists("metrics.json"):
            with open("metrics.json", "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        print(f"Unable to load metrics.json: {e}", file=sys.stderr)
    return {}

def main():
    repo, pr_number, head_sha, _ = github_env()
    pr_files = list_pr_files(repo, pr_number)

    ipynb_results = []
    data_files = []

    for f in pr_files:
        status = f.get("status")  # added/modified/removed/renamed
        if status in ("removed",):
            continue
        path = f.get("filename", "")
        lowered = path.lower()

        if lowered.endswith(IPYNB_EXT):
            content = fetch_file_content(repo, path, head_sha)
            has_outputs, note = notebook_has_outputs(content)
            if has_outputs:
                ipynb_results.append((path, "⚠️ outputs present"))
            else:
                ipynb_results.append((path, note))
        else:
            for ext in DATA_EXTS:
                if lowered.endswith(ext):
                    data_files.append(path)
                    break

    metrics = load_metrics_if_any()
    body = build_comment(repo, pr_number, head_sha, ipynb_results, data_files, metrics)
    existing = find_existing_comment(repo, pr_number)
    post_or_update_comment(repo, pr_number, body, existing)

if __name__ == "__main__":
    try:
        main()
    except requests.HTTPError as e:
        print(f"GitHub API error: {e} - {getattr(e, 'response', None) and getattr(e.response, 'text', '')}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Unhandled error: {e}", file=sys.stderr)
        sys.exit(1)
