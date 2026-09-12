#!/usr/bin/env bash
#
# Publish QuickLook Office 极速预览版 (Turbo) to GitHub under your own account.
#
#   bash tools/publish_github.sh --mode repo          # standalone repo with just this package (default path)
#   bash tools/publish_github.sh --mode repo --dry-run
#   bash tools/publish_github.sh --mode repo --name quicklook-officeviewer-turbo
#   bash tools/publish_github.sh --mode fork          # (optional) fork QL-Win/QuickLook + a branch on it
#
# Requirements:
#   * gh (GitHub CLI) installed AND authenticated:  gh auth login
#   * push access to github.com/<you>/<repo>.git
#
# The standalone repo is the recommended shape for this project: users searching for
# "QuickLook Office preview slow" / "Protected View popup" need to find THIS repo, and a
# standalone repo is far easier to discover than a branch buried inside a fork.
# The upstream project (QL-Win/QuickLook) is credited prominently in the README, and
# nothing is ever pushed to it.
#
set -euo pipefail

MODE=""
DRY=0
FORK_NAME="QuickLook"
STANDALONE_NAME="quicklook-officeviewer-turbo"
BRANCH="office-preview-patched"
UPSTREAM="QL-Win/QuickLook"
VERSION="4.5.0-turbo.3"
TAG="v4.5.0-turbo.3"
# ≤20 topics, lowercase, hyphens only. These are how people find this repo.
TOPICS="quicklook quicklook-plugin quicklook-officeviewer office-preview preview-handler file-preview protected-view zone-identifier performance word-preview excel-preview powerpoint-preview dotnet wpf windows"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --mode)   MODE="${2:-}"; shift 2 ;;
    --dry-run) DRY=1; shift ;;
    --name)   STANDALONE_NAME="${2:-}"; shift 2 ;;
    --fork-name) FORK_NAME="${2:-}"; shift 2 ;;
    --branch) BRANCH="${2:-}"; shift 2 ;;
    -h|--help) sed -n '2,21p' "$0"; exit 0 ;;
    *) echo "unknown argument: $1" >&2; exit 2 ;;
  esac
done

# standalone repo is the default mode
MODE="${MODE:-repo}"
if [[ "$MODE" != "fork" && "$MODE" != "repo" ]]; then
  echo "usage: bash tools/publish_github.sh [--mode repo|fork] [--dry-run]" >&2
  exit 2
fi

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
REPO="$ROOT/repo"
PKG="$ROOT/dist/QuickLook-OfficeViewer-Turbo"
GH="/c/Program Files/GitHub CLI/gh.exe"
[[ -x "$GH" ]] || GH="$(command -v gh || true)"
if [[ -z "$GH" || ! -x "$GH" ]]; then
  echo "ERROR: gh (GitHub CLI) not found. Install it, then run: gh auth login" >&2
  exit 1
fi

run() {
  echo "  \$ $*"
  if [[ $DRY -eq 1 ]]; then return 0; fi
  "$@"
}

# gh.exe is a native Windows binary and does its own glob matching on the asset arguments of
# `gh release create/upload`. Handing it an MSYS-style "/c/..." path makes it report
# "no matches found for ..." instead of uploading. Convert to mixed-mode "C:/..." which both
# gh and Windows understand, and which contains no glob metacharacters.
winpath() {
  if command -v cygpath >/dev/null 2>&1; then
    cygpath -m "$1"
  else
    printf '%s' "$1"
  fi
}

# Run git inside a directory. NOTE: this machine's git is a MinGW build
# (/mingw64/bin/git) that does NOT translate MSYS-style "/c/..." paths, so
# `git -C /c/...` fails with "cannot change to ...". Always `cd` instead.
gitdir() {
  local dir="$1"; shift
  echo "  \$ (cd $dir && git $*)"
  if [[ $DRY -eq 1 ]]; then return 0; fi
  ( cd "$dir" && git "$@" )
}

DRYLABEL=""
[[ $DRY -eq 1 ]] && DRYLABEL=", dry-run"
echo "=== publish to GitHub (mode: $MODE$DRYLABEL) ==="

# --- 0. preflight -----------------------------------------------------------
if ! "$GH" auth status >/dev/null 2>&1; then
  cat >&2 <<'EOF'
ERROR: gh is not authenticated.

  Run this once, in your own terminal (it opens a browser):

      gh auth login

  If gh is not on PATH, use the full path:

      "/c/Program Files/GitHub CLI/gh.exe" auth login

  Then re-run this script.
EOF
  exit 1
fi

USER="$("$GH" api user --jq .login)"
echo "authenticated as : $USER"
echo "package          : $PKG"
[[ -f "$PKG/MANIFEST.txt" ]] || { echo "ERROR: package missing - run: python tools/make_package.py" >&2; exit 1; }

# --- mode: standalone repo --------------------------------------------------
if [[ "$MODE" == "repo" ]]; then
  WORK="$ROOT/dist/github-backup"
  echo
  echo "[1/3] building a standalone repository in $WORK"
  run rm -rf "$WORK"
  run cp -r "$PKG" "$WORK"
  gitdir "$WORK" init -q -b main
  gitdir "$WORK" add -A
  if [[ $DRY -eq 0 ]]; then
    ( cd "$WORK" && git -c user.name="$(git config --get user.name 2>/dev/null || echo "$USER")" \
                        -c user.email="$(git config --get user.email 2>/dev/null || echo "$USER@users.noreply.github.com")" \
                        commit -q -m "QuickLook Office 极速预览版 $VERSION

Non-official plugin-level patch for QL-Win/QuickLook 4.5.0 that solves:

  1. Office preview performance - keeps the Office preview-handler class factory alive
     so consecutive previews of Word/Excel/PowerPoint files hit a ~2 ms warm activation
     instead of restarting the Office process (482-863 ms) every single time, and adds an
     idle watchdog so the Office processes are NOT kept resident forever (was ~355 MB).
  2. Protected View - documents carrying a Zone.Identifier are previewed via an unblocked
     temporary copy, so the user's original file is never modified; a notice strip explains
     this and offers an explicit button to unblock the source file. The old blocking Yes/No
     confirmation dialog is removed.
  3. Bilingual UI - the notice strip follows the Windows UI language (Chinese/English),
     forceable via the UiLanguage setting.

Install: unzip, double-click install.bat. rollback.bat restores stock.
Only ONE file is deployed: QuickLook.Plugin.OfficeViewer.dll.

Upstream: https://github.com/QL-Win/QuickLook (GPL-3.0). Unofficial; upstream is not
affiliated with nor does it endorse this patch." )
  else
    echo "  \$ (cd $WORK && git commit ...)"
  fi
  echo
  echo "[2/3] creating $USER/$STANDALONE_NAME and pushing"
  # Create the repo WITHOUT --source: gh is a native Windows binary and would not
  # understand an MSYS-style "/c/..." path. We add the remote and push with git
  # ourselves (over SSH, which needs no credential helper).
  if "$GH" repo view "$USER/$STANDALONE_NAME" >/dev/null 2>&1; then
    echo "  repo already exists: $USER/$STANDALONE_NAME"
  else
    run "$GH" repo create "$STANDALONE_NAME" --public \
        --description "QuickLook Office 极速预览版 (Turbo) - non-official plugin patch: instant consecutive Word/Excel/PowerPoint previews (~2ms warm activation instead of restarting Office every time) + Protected View documents previewed as an unblocked copy. Based on QL-Win/QuickLook (GPL-3.0)."
  fi
  gitdir "$WORK" remote remove origin 2>/dev/null || true
  gitdir "$WORK" remote add origin "git@github.com:$USER/$STANDALONE_NAME.git"
  # --force: this repo contains nothing but our own package, so re-publishing must be
  # idempotent (each run builds a fresh single-commit history). No user content is at risk.
  gitdir "$WORK" push -u --force origin main
  echo
  echo "[3/4] setting discovery topics"
  if [[ $DRY -eq 0 ]]; then
    for t in $TOPICS; do
      "$GH" repo edit "$USER/$STANDALONE_NAME" --add-topic "$t" >/dev/null
    done
  fi

  # --- release --------------------------------------------------------------
  # The Release is where people actually download from, so it must carry real assets:
  #   * the bare plugin DLL  - the whole patch is one file, so this is a valid download
  #   * the plugin-only zip  - the same DLL plus install.bat / rollback.bat
  #   * the full package zip - source, patch, docs, tools
  echo
  echo "[4/4] publishing the GitHub Release $TAG"
  NOTES="$ROOT/packaging/RELEASE-NOTES.md"
  [[ -f "$NOTES" ]] || { echo "ERROR: missing $NOTES" >&2; exit 1; }
  DLL="$PKG/plugin/QuickLook.Plugin.OfficeViewer.dll"
  PLUGZIP="$ROOT/dist/QuickLook-OfficeViewer-Turbo-plugin-only.zip"
  FULLZIP="$ROOT/dist/QuickLook-OfficeViewer-Turbo-$VERSION.zip"
  for f in "$DLL" "$PLUGZIP" "$FULLZIP"; do
    [[ -f "$f" ]] || { echo "ERROR: missing release asset: $f" >&2; exit 1; }
  done
  DLL_W="$(winpath "$DLL")"
  PLUGZIP_W="$(winpath "$PLUGZIP")"
  FULLZIP_W="$(winpath "$FULLZIP")"
  NOTES_W="$(winpath "$NOTES")"
  TITLE="QuickLook Office 极速预览版 $VERSION / Turbo"
  if "$GH" release view "$TAG" --repo "$USER/$STANDALONE_NAME" >/dev/null 2>&1; then
    echo "  release $TAG already exists - updating"
    run "$GH" release edit "$TAG" --repo "$USER/$STANDALONE_NAME" --title "$TITLE" --notes-file "$NOTES_W"
    run "$GH" release upload "$TAG" --repo "$USER/$STANDALONE_NAME" --clobber "$DLL_W" "$PLUGZIP_W" "$FULLZIP_W"
  else
    # This repo contains nothing but our own package and every publish force-pushes a fresh
    # single-commit history, so a tag left over from an earlier run would pin the Release to a
    # commit that no longer exists on main. Drop it so the tag is recreated at the new HEAD.
    gitdir "$WORK" push origin ":refs/tags/$TAG" 2>/dev/null || true
    run "$GH" release create "$TAG" --repo "$USER/$STANDALONE_NAME" \
        --title "$TITLE" --notes-file "$NOTES_W" \
        "$DLL_W" "$PLUGZIP_W" "$FULLZIP_W"
  fi

  if [[ $DRY -eq 0 ]]; then
    URL="$("$GH" repo view "$USER/$STANDALONE_NAME" --json url --jq .url)"
    echo
    echo "done: $URL"
    echo "release: $URL/releases/tag/$TAG"
    echo "topics: $TOPICS"
  fi
  exit 0
fi

# --- mode: fork -------------------------------------------------------------
echo
echo "[1/5] forking $UPSTREAM (no-op if the fork already exists)"
if "$GH" repo view "$USER/$FORK_NAME" >/dev/null 2>&1; then
  echo "  fork already exists: $USER/$FORK_NAME"
else
  run "$GH" repo fork "$UPSTREAM" --clone=false --remote=false
fi

echo
echo "[2/5] making sure the local clone has full history (needed to push)"
if [[ -f "$REPO/.git/shallow" ]]; then
  echo "  shallow clone detected - fetching full history (GitHub rejects pushes from shallow clones)"
  gitdir "$REPO" fetch --unshallow origin
else
  echo "  already a full clone."
fi

echo
echo "[3/5] staging the portable package inside the branch"
DOCDIR="$REPO/docs/quicklook-officeviewer-turbo"
[[ -f "$DOCDIR/MANIFEST.txt" ]] && run rm -rf "$DOCDIR"
run mkdir -p "$DOCDIR"
run cp -r "$PKG/." "$DOCDIR/"
gitdir "$REPO" add docs/quicklook-officeviewer-turbo
if [[ $DRY -eq 0 ]]; then
  if ( cd "$REPO" && git diff --cached --quiet ); then
    echo "  package already committed on this branch - nothing to add."
  else
    ( cd "$REPO" && git commit -q -m "docs: ship the QuickLook Office 极速预览版 package

Self-contained bundle for the change: patched source, git patch vs upstream a3ab193,
prebuilt plugin DLL, official stock DLL for rollback, install/rollback scripts,
verification tools and analysis docs." )
  fi
fi

echo
echo "[4/5] pushing branch '$BRANCH' to your fork"
gitdir "$REPO" remote remove fork 2>/dev/null || true
gitdir "$REPO" remote add fork "git@github.com:$USER/$FORK_NAME.git"
gitdir "$REPO" push -u fork "$BRANCH"
gitdir "$REPO" tag -f "$TAG" "$BRANCH"
gitdir "$REPO" push -f fork "refs/tags/$TAG"

echo
echo "[5/5] result"
if [[ $DRY -eq 0 ]]; then
  URL="$("$GH" repo view "$USER/$FORK_NAME" --json url --jq .url)"
  echo "  fork   : $URL"
  echo "  branch : $BRANCH"
  echo "  tree   : $URL/tree/$BRANCH"
  echo
  echo "  Upstream ($UPSTREAM) was NOT touched, and a fork cannot write to it."
fi
