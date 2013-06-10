DataImport
==========

Generating Excel Templates and Importing the Data - Using Spring MVC 3.2.2, Hibernate, Apache POI and Maven.

NOTES
=========
- Keep this as lightweight as possible, no need to add frameworks unless it solves complexity.
- Only Servlet (no JSP), keep everything (js and css) in index.html only
- Validate mime type of file on client side if possible
- Use POI only for reading, no need to generate templates using POI, keep the templates as resources within code