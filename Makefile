.PHONY: pages
pages:
	npm run build
	cp CNAME docs/
	git restore docs/CNAME docs/google\*

.PHONY: serve
serve:
	cd docs; python -mSimpleHTTPServer; cd -
