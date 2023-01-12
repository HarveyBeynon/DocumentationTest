# DocumentationTest
This GitHub repo downloads pages as .doc files from a given confluence space. The script then orders the page alphabetically, and converts the files to .rst. The repo is then pushed to GitHub which is connected to a [Read The Docs project](https://harveybeynon-testdocs.readthedocs.io/en/latest/index.html) via a webhook. 

Within the docs directory is an index.rst page, which acts as the homepage for the docs. This page includes a table of contents for the remaining pages. These pages are found within the docs/pages directory. These pages are list alphabetically - **hence the page titles need to be numbered in order like so:**
- 1.1 Page title (Section 1 page 1)
- 1.2 Page title (Section 1 page 2)
- 2.1 Page title (Section 1 page 1)

Images for the documentation need to be saved to the image folder in the same repo as the confluence_to_read_the_docs script.