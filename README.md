# VariableVal
This program will allow a user to enter in a date range to receive specific data from the Inventory Valuation page on Plex.

Versions of python and installed modules:
- python 3.7.8
- selenium 3.141.0
- ChromeDriver 80.0.3987.106
- Visual Studio 16.8.3
- numpy 1.19.3 (the 1.19.4 update messes with selenium)
- openpyxl 3.0.5

Requirements:
- Wanco Plex account with access to the Inventory Valuation page

The peculiar thing about the Inventory Valuation page is that it takes in two dates at once, but these dates don't represent a range.
Instead, each date represents some data at that instance. These two dates are inventory date and cost date.

There are a few settings to adjust, such as Group by "Building" and ensuring that only "Parts" is check-marked in the bottom row
of stuff. It's optional to uncheck the box next to "Customer" since whatever data this could provide isn't counted anyway.
These adjustments are constant and will always be activated for every search. The dates, however, are not. They're set by the
user.

Since the search results only show data for one particualr instance in time, the program will have to search multiple times for 
the (very common) cases where the date range is more than a day. That being said, I think the program could allow the user to
enter in the dates in any order. Besides the exception of both dates being the same, one date will always be farther back in
time compared to the other. The dates will have to be compared numerically to figure out the days between them anyway, might
as well implement the "any order" feature. 
