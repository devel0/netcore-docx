#!/bin/bash

exdir="$(dirname `readlink -f "$0"`)"

DOCSDIR="$exdir/docs"

rm -fr "$DOCSDIR"

mkdir "$DOCSDIR"

cd "$exdir"

doxygen

rsync -arvx "$exdir/test/" "$DOCSDIR/test/" \
    --exclude=bin \
    --exclude=obj