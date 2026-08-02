#!/usr/bin/env bash

set -euo pipefail

if [ "$#" -ne 2 ] || [ -z "$1" ] || [ -z "$2" ]; then
  echo "usage: rclone-read-metadata.sh <remote-object> <destination>" >&2
  exit 64
fi

readonly source_object="$1"
readonly destination="$2"
readonly temporary="${destination}.partial"

cleanup() {
  rm -f "${temporary}"
}
trap cleanup EXIT

for attempt in 1 2 3; do
  if timeout 60s rclone cat "${source_object}" \
    --retries 2 \
    --low-level-retries 3 \
    --retries-sleep 2s > "${temporary}"; then
    mv "${temporary}" "${destination}"
    trap - EXIT
    exit 0
  fi
  rm -f "${temporary}"
  echo "::warning::rclone metadata read attempt ${attempt} failed"
  if [ "${attempt}" = "3" ]; then
    exit 1
  fi
  sleep "$((attempt * 3))"
done
