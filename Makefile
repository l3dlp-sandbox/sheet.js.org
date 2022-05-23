.PHONY: pages
pages:
	npm run build
	cp CNAME docs/
