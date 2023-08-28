# powerbi-webscraper
A webscraping program that gather info from the last months updates to Power BI, and adds the changes to an excel an creates a power point presentation. Final assigment in my Noroff Python course


******************************    READ ME   ******************************


NAME: POWERBI-BLOGG-SCRAPER
VERSION.NR 1.1
LAST UPDATE: 21.06.2023
CREADTED BY: BENDIK MENDES DAHL



****************************** INTRODUCTION ******************************

Thank you for taking the time to review this great program.

This program is developed to automate the information gathering process 
about the latest news and updates about Microsoft Power BI- and put the 
updates into a powerpoint presentation. 

The program is developed to help my own workflow, and the only intended 
user of the program itself is myself. Maybe once a quarter I will make 
a presentation about the last changes, and instead of manually gathering
the updates from this blogg: https://powerbi.microsoft.com/en-us/blog/
this program will do it for me - utilizing a process called web-scraping



****************************** HOW TO RUN IT ******************************

Since this program is not designed to be used by other people (at least not
other non-technical people) it requires jupyter notebook to be used.


1. open the file in jupyter notebook
2. on line 6 in the first cell - select how many pages of the blogg you want
   to include (the number of publications varies from month to month, but 
   it is more or like 1 page per month - so select 3 if you want to get an
   overivew for more or less the last three months.
3. run the first cell
4. open the excel file that has now been created in the folder. Verify that
   you have selected the desiered time frame, and verify the data
5. makes sure the powerpoint file "template1" is the same folder as the rest
   of the program
5. run the second cell 
6. now your powerpoint presentation has been generated!
7. save both the excel file and the powerpoint somewhere else
8. feel free to edit the slides to your liking
9. enjoy!

******************************  KEEP IN MIND  ******************************

---WEBSCRAPING---

The chances for something will break over the next quarters are pretty high
The program refers to specific HTML attributes/names/objects on the blog,
which Microsoft could change every now and then

So when generating the excel file, see if all the info you want is there,
and feel free to cross reference it with the actual blog and the posts.


---FEATURED POST---

The code should not list/referance the same blogpost more than once, but at
the same time you would not expect Microsoft to post the same post twice...
However:
On the top of the blog there is a "Featured post" which Microsoft think is
the most important news at the moment (not the newest). 
When going through the different pages of the blog, this Featured post is
always on the top. 

While the program is designed to take this Featured post into account, 
keep an extra eye on this going forward. Should there be any issues with the
HTML code this could be a good place to start


---CLOSE THE FILES ---

Remember to not have the excel file or the powerpoint files open when running 
the code.

The first time this will most likely not be an issue, but be aware when 
debugging or developing/changing the code that this can make the program to 
crash. No big deal, just close the program and try again! 


---POWERPOINT TEAMPLATE---

If you wish to update the powerpoint template - be careful! 
On the top meny: view - Slide Master -> make the changes you want -> save

see more here: "https://python.plainenglish.io/turn-your-excel-sheets-into-
a-powerpoint-presentation-just-in-a-couple-of-minutes-62ae4b8bc13b"


---OVERRIDING OUTPUT FILES---
Since there might be some debugging needed, or try and error with selecting
the right number of pages relevant for this month/quarter's presentation
the program is designed to override the exsisting excel and powerpoint files
if they have already been generated. so remember to 'save as' these files 
somewhere else after you have run through all the steps - So you don't lose 
them the next time you run the program


---CHANGE LOG---
Due to the sporadic and pragmatical use of the program, there is not designed
a nice log og which blogpost have been fetched by the program and put into the 
excel and ppt earlier.

So if you want to keep track of which posts you have included earlier, 
remember to save the excel file, and cross check manually by looking at the
date of the last post in your last presentation (remember the Featured post
can have a date that differs from the chronological order)


---SUPPORT---

You have build it! Now let's see if you manage to support it!

24/7 Support: +47 815 493 00



 
******************************   ROAD MAP   ******************************

The next things I would have liked to built is:

-fetching the pictures from the relevant blogpost, and included them 
 to the powerpoint presentation
-some kind of change log/version control of what has been fetched before
-making the powerpoint template more customizable
-better user input validation
-building a library of all the blogposts
-split the categories string into each category  

and in the longer run:
-fetch news/blogs from ther soures 
-link with Chat GPT or similar to read the blog and be able to summarize
 (eg. in bullet points, 50 or 200 words, translate to Norwegian ect)



******************************    THANKS    ******************************
