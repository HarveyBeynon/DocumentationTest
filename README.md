# DocumentationTest
This GitHub repo downloads pages as .doc files from a given confluence space. The script then orders the page alphabetically, and converts the files to .rst. The repo is then pushed to GitHub which is connected to a [Read The Docs project](https://harveybeynon-testdocs.readthedocs.io/en/latest/index.html) via a webhook. 

Within the docs directory is an index.rst page, which acts as the homepage for the docs. This page includes a table of contents for the remaining pages. These pages are found within the docs/pages directory. These pages are list alphabetically - **hence the page titles need to be numbered in order like so:**
- 1.1 Page title (Section 1 page 1)
- 1.2 Page title (Section 1 page 2)
- 2.1 Page title (Section 1 page 1)

Images are appended to the pages in order by replacing the image directive in the rst files with the contents of the docs/pages/media directory. **Images in this directory are listed alphabetically** therefore all images need to be named in order. **For best practice, name the images with the same number as thier corresponding pages like so**
- 1.1.1 image.jpg (First image in section 1 page 1)
- 1.1.2_image.jpg (Second image in section 1 page 1)
- 1.2.1_image.jpg (First image in section 1 page 2)
- 1.2.2_image.jpg (Second image in section 1 page 2)
- 2.1.1_image.jpg (First image in section 2 page 1)

**Note that any images used on the confluence page needs to be saved it this repo's docs/pages/media directory**