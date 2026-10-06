#!/bin/bash
# sync-mesh.sh — Materialize remote squad state locally
#
# Usage: ./sync-mesh.sh [path-to-mesh.json]
#        ./sync-mesh.sh --init [path-to-mesh.json]
# Requires: jq, git, curl

set -euo pipefail

if [ "${1:-}" = "--init" ]; then
  MESH_JSON="${2:-mesh.json}"
  if [ ! -f "$MESH_JSON" ]; then
    echo "$MESH_JSON not found" >&2
    exit 1
  fi

  squads=$(jq -r '.squads | keys[]' "$MESH_JSON")
  for squad in $squads; do
    mkdir -p "$squad"
    if [ ! -f "$squad/SUMMARY.md" ]; then
      printf '# %s\n\n_No state published yet._\n' "$squad" > "$squad/SUMMARY.md"
    fi
  done

  if [ ! -f "README.md" ]; then
    {
      echo "# Squad Mesh State Repository"
      echo ""
      echo "This repository tracks published state from participating squads."
      echo ""
      echo "## Participating Squads"
      echo ""
      for squad in $squads; do
        zone=$(jq -r --arg squad "$squad" '.squads[$squad].zone' "$MESH_JSON")
        echo "- **$squad** (Zone: $zone)"
      done
      echo ""
      echo 'Each squad directory contains a `SUMMARY.md` with its latest published state.'
      echo 'State is synchronized using `sync-mesh.sh` or `sync-mesh.ps1`.'
    } > README.md
  fi

  echo "Mesh state repository initialized"
  exit 0
fi

MESH_JSON="${1:-mesh.json}"

for squad in $(jq -r '.squads | to_entries[] | select(.value.zone == "remote-trusted") | .key' "$MESH_JSON"); do
  source=$(jq -r --arg squad "$squad" '.squads[$squad].source' "$MESH_JSON")
  ref=$(jq -r --arg squad "$squad" '.squads[$squad].ref // "main"' "$MESH_JSON")
  target=$(jq -r --arg squad "$squad" '.squads[$squad].sync_to' "$MESH_JSON")

  if [ -d "$target/.git" ]; then
    git -C "$target" pull --rebase --quiet
  else
    mkdir -p "$(dirname "$target")"
    git clone --quiet --depth 1 --branch "$ref" -- "$source" "$target"
  fi
done

for squad in $(jq -r '.squads | to_entries[] | select(.value.zone == "remote-opaque") | .key' "$MESH_JSON"); do
  source=$(jq -r --arg squad "$squad" '.squads[$squad].source' "$MESH_JSON")
  target=$(jq -r --arg squad "$squad" '.squads[$squad].sync_to' "$MESH_JSON")
  auth=$(jq -r --arg squad "$squad" '.squads[$squad].auth // ""' "$MESH_JSON")

  mkdir -p "$target"
  auth_args=()
  if [ "$auth" = "bearer" ]; then
    token_var="$(echo "${squad}" | tr '[:lower:]-' '[:upper:]_')_TOKEN"
    if [ -z "${!token_var:-}" ]; then
      echo "Missing bearer token for $squad" >&2
      exit 1
    fi
    auth_args=(--header "Authorization: Bearer ${!token_var}")
  fi

  curl --silent --show-error --fail "${auth_args[@]}" -o "$target/SUMMARY.md" -- "$source"
done

echo "Mesh sync complete"
