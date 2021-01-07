# VariableVal
This program will allow a user to enter in a date range to receive specific data from the Inventory Valuation at Standard Cost page on Plex.

Versions of python and installed modules:
- python 3.7.8
- selenium 3.141.0
- ChromeDriver 80.0.3987.106
- Visual Studio 16.8.3
- numpy 1.19.3 (the 1.19.4 update messes with selenium)
- openpyxl 3.0.5

Requirements:
- Wanco Plex account with access to the Inventory Valuation page

Update 1/7/2021: The date range feature was successfully implemented. Although the program prompts the user to enter the "first
date" before the "second date", they can actually input them in any order and the program will recognize which is the oldest.
There was also an issue with Plex concerning building caches. The only dates I could access with the time set to 5:00 PM were 
1/1/2021, 1/2/2021, 1/3/2021, and 1/4/2021. When I attempted to access any other dates, notably the ones from last year, no
results would appear and red text came up saying "The result you're trying to access with this filter doesn't have a cache built.
A background process has been started. The process' completion will be sent to your account's registered email." If the page is
refreshed or the search button is clicked again, then the results will show up. Next, my email's inbox contains an email telling
me "Inventory Valuation may now be run with the following filters" and below that are the dates and times the program entered
for that one search. There was also a strange case for 11/1/2020. When using this date, Plex would take a bit to load a new page.
This is usually a sign that Plex is busy trying to pull up result. However, what came up wasn't results, but rather no results.
This is odd because given the nature of the data, it seemed that every day should have results. Another thing that stood out was
that the date now said 10/31/2020 instead of 11/1/2020. What exactly is the cause of all of these weird occurrences? Regardless,
they shouldn't affect the program's operation that much. The first issue just needs a page refresh and the second one gets
skipped anyway since no results turn up.

The peculiar thing about the Inventory Valuation page is that it takes in two dates at once, but these dates don't represent a range.
Instead, each date represents some data at that instance. These two dates are inventory date and cost date. For this program's purpose,
there's no need to put in different dates for a single search. Each search will utilize the same dates.

There are a few other settings to adjust, such as Group by "Building" and ensuring that only "Parts" is check-marked in the bottom row
of stuff. It's optional to uncheck the box next to "Customer" since whatever data this could provide isn't counted anyway.
These adjustments are constant and will always be activated for every search. The dates, however, are not. They're set by the
user.

Since the search results only show data for one particular instance in time, the program will have to view multiple days for 
the (very common) cases where the date range is more than a day. That being said, I think the program could allow the user to
enter in the dates in any order. Besides the exception of both dates being the same, one date will always be farther back in
time compared to the other. The dates will have to be compared numerically to figure out the days between them anyway, might
as well implement the "any order" feature. 

The final excel sheet consists of all the final rows of the searched-up pages. The only difference is that the "Vendor Managed"
column is replace by a date column to signify when the data occurred. 
