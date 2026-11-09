# Reference tools (LibreOffice, openpyxl, Apache POI) are optional.
# diffcheck skips any engine that is not installed.
# POI jars default to /tmp/poi-jars; override with POI_JAR_DIR.

.PHONY: test diffcheck xlsxdump

test:
	go test ./...

xlsxdump:
	go build -o $(CURDIR)/xlsxdump ./cmd/xlsxdump

diffcheck: xlsxdump
	python3 scripts/diffharness.py --roundtrip testdata/excel
