# ComponentFinder
This repository contains a simple Python script for finding the cheapest website for purchasing a certain electronic component. The script checks sites like Mouser, RS-Online, Farnell, and Digikey, and returns all prices in an Excel sheet, where the cheapest can easily be found. Adding more sites is easy, and so is customizing the output. Eventually, a UI will be created for ease of interacting with the program.

## Known Bugs
There are still some bugs left to be ironed out due to the flexibility of the program.

- If the web driver goes directly into the product page on RS Online, it cannot correctly find the component name.

- The Farnell description occasionally returns extra text like "You have previously bought this product". This may also be an issue with Farnell.

- If the product is not found, websites may show similar results, returning an incorrect product.

## Future Work
Some functionality and restructuring is still needed for ease-of-use. The list is of descending priority.

- Also include prices when buying in bulks (10, 25, 100 components at a time); may be possible to implement on Mouser by finding <tr data-qty=x> and then finding <span id=lblPrice>; on RS by looking for Price-Break-Quantity and then Price-Break-Price; on Farnell by checking quantity and price.

- Make sure that all currencies are in euros to ease the comparision.

- Returning the actual component link instead of just the search term for RS Online.

- A more sophisticated decision-making than simply picking the cheapest, as it occasionally results in errors.

- Create a GUI for the software, for example using Qt.

- Allow loading a file (csv or new-line separated) into the program and search for each component automatically.

- Use APIs for the websites instead of scraping the HTML, for example with JSON. Both Farnell and Mouser have an API, RS Online might too. This would also allow adding Digi-Key, as they provide a search API.