# Local Wiki
A local wikipedia like site which converts word document to website pages and makes them searchable in an instant

## index.html
An edited version of the wikipeida template that can be found: https://html5-templates.com/preview/wikipedia.html

## script.js
This impliments a contents page automatically based on the contents of the site.
It adds every h3 element to the contents list and makes a link to it when clicked on in the contents page to take a user straight to it.

## style.css
The style required to make the website work.
Currently the file has an ugly search field drop down as this is in development.

## search.js
Simple search where by it looks through the contents of the site to make a JSON object of all the possible things it could be and returns the results.
Going forward this will also look at the contents of the file which is a function being worked on in the Compile-Files.ps1 script.
The JSON object is currently made upon running the site, this will be improved to a pre-made local JSON object closer to the time.
