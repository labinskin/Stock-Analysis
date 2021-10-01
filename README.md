# Green Stocks Analysis

By Nicholas Labinski

### Background

Steve had asked me to write VBA code to help him analyze green stocks that his parents have invested in, as well as other potential green stocks. I initially analyzed the DQ stock for Steve. However, after this analysis, he wanted to be able to analyze the total volume and return percentage of other green stocks. I have written code to help him analyze both DQ and other stocks, and he has been appreciative of the work I have done for him. However, he wanted to be able to analyze the entire dataset, and possibly larger datasets in the future, with the click of a button. Because of this expansion in size of the dataset(s), my code will have to be updated to work as quickly and efficiently as possible, as the current code might not be able to handle the volume of data. The solution to this problem is to refactor my original code, making it more efficient and taking less time to run, especially with an expanded dataset.

### Results

**Stock Performance**

Insert 2017 Picture

The return price for 2017 based on the refactored code analysis shows that all the stocks, except TERP, had a positive return, with four stocks (DQ, ENPH, FSLR, and SEDG) reaching triple digit returns. However, the 2018 returns tell a different story (see image below)

Insert 2018 Picture

Only two stocks had positive returns (ENPH and RUN). All other stocks had a negative return. Based on the returns, green stocks appear to be highly volatile. Only ENPH had positive returns in both years. Based on the returns, I would warn Steve and his parents to be wary of investing in green stocks in general. However, if they really wanted to invest in a green stock, I'd recommend ENPH, as it has consistently performed well and brought back positive returns.

**Code Performance**

Below are two images containing the times for the original code.

Insert Image 2017

Insert Image 2018

Below are the two images containing the times for the refactored code.

Insert Image 2017

Insert Image 2018

For 2017, the code ran .711 seconds faster, while for 2018, the code ran .6953 seconds faster. While these are only partial seconds faster, Steve would like this code to work on larger datasets. The increase in volume of data could potentially make the difference between the original code and refactored code larger. Now to examine how these differences came about, we are going to compare the two sets of code. First, below immediately you will see the major sections of first the original code, then the refactored code.

Insert Image Original

Insert Image Refactored

One of the major differences is that in the original code, there is a nested loop, whereas in the refactored code, the nested loop is no longer present. This makes less complicated code overall, as writing a nested loop takes time and precision. The main section and differences I want to touch on concern 5 A-C of the original code and 3 A-C of the refactored code, as it is in these sections of code that major differences appear.

![](https://github.com/labinskin/Stock-Analysis/blob/main/Original%20Code%205%20A_C.png)

Insert Image 3



### Summary

**Benefits of Refactoring In General**

The benefits of refactoring code are many. In this particular case, the code became simpler and more efficient, allowing it to run faster. For 2017, the code ran .711 seconds faster (see images below). For 2018, the code ran .6953 seconds faster (see images below).

Insert Images

While these mere fractions of a second don't seem like much for this current dataset, the differences would likely be larger with expanded and larger datasets. These minor improvements can add up, creating an application and program that has better performance. Additional benefits of refactoring code include keeping it clean. While this set of code was written by myself, larger projects usually have more than one code writer. The more writers, the more potential for repetition and redundancies, and/or unnecessary variables and loops, to name just a few potential pitfalls. Any of these could make the program inefficient, but by refactoring it, these problems can be avoided. Another benefit, while not necessarily part of this project, is that code can become outdated and in need of updating. By refactoring it to the latest updates, the coder can improve performance of the application.

**Drawbacks of Refactoring In General**

While there are many benefits to refactoring code, there are a few potential drawbacks that need to be mentioned. First, refactoring can take time. This, like most projects, had a hard deadline. While this was a fairly simple code and project, larger projects that are either planning on launching into the market or public as quickly as possible, might want to wait on refactoring their code. This project may launch with some bugs that refactoring will fix in the future, but the project may need to sent out for investors to validate their support. Fortunately, for this project, refactoring was the main objective.

Tied closely to the time argument is possibly not figuring out how to properly refactor the code (or in some cases, having old code that is a complete mess and not knowing where to begin). As an individual new to coding and data analytics, working through the different ways to shorten and improve the code is a process that takes time. And sometimes, my refactored code ran into errors when I tested it, whether it was forgetting to add a variable somewhere or figuring out if there should be nested loop in the process, it took me time to work out these errors. And that time, in a real job situation, might not have been there. The original code, while less efficient, might have had to go for the launch of the project, with time set aside in the future for refactoring because the product needed to be usable.
