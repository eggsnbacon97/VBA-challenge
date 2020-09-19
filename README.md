# VBA-challenge
## VBA script for analyzing stock data. Here's what it does!

- Creates the ticker symbol.

- Pulls the yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

- Pulls the percent change from opening price at the beginning of a given year to the closing price at the end of that year.

- Pulls the total stock volume of the stock.

- Formats postive/negative yearly change results.

- Gives the greatest percent increase/decrease for that year, as well as the greatest total volume for that year.

## Here's how it works!

### 1. Does a lot of pretty VBA formatting (bolding, percent formatting, coloring) with values like:
```
.Font.Bold=true
.NumberFormat="0.00%"
.Interior.ColorIndex = 4
```
This is pretty much setting your formatting up for ease of reading. There's no point to having all that data if you can't read it!
As you may notice, there are two "Ticker" values, and yes, one does reference the other by using:
```
ws.Range("P1").Value = ws.Range("I1").Value
```
### 2. Declares all the variables (well, *almost*)
Here, I am declaring variables that will be used later on in the script. Stuff like *Ticker_Symbol*, *Yearly_Change*, etc. It's all going to come in handy later!
```
    Dim Ticker_Symbol As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim LastRow As Long
    Dim Total_Stock_Volume As Double
    Dim Summary_Table_Row As Long
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Old_Amount As Long
    Dim i As Long
```
So now that I've declared those variables, we only need to set a few. The reason behind this is because the default value for an empty variable is zero. 

### 3. Start the logic!
Now the script actually begins to "do it's thing". We're going to give it a loop to run, a bunch of conditions, and let it run!
I won't go through all the code now (you can look at it in the .bas file), but here are the conditions it's given:

- If the ticker symbol found beneath the cell we're looking in is different, print it! And then:
- Add up all of our specific stock volume for a total, and print that. Then, reset the total back to zero.
- Take the difference of our open and close price, then calculate and print the yearly change.
- Divide the yearly change over the open price to get our percent change, and print that thing out!
- If the yearly change amount is greater than or less than zero, format accordingly.
- All done, go to the next row!

This will run until it hits the last row, which we defined earlier.

### 4. Greatest Increase, Decrease, and Greatest Total Volume (the extra bit)
Now we are going to make another loop (not a nested one). This one will run through our Percent Change and Total Stock Volume columns, and get our minimum and maximums.
Here are the conditions given:

- If the cell in the range of the percent change column is the greatest, print it out accordingly.
- If the cell in the range of the percent change column is the smallest, print it out accordingly.
- If the cell in the range of the total stock volume column is the greatest, print it out accordingly.

For all of those conditions, there's also a "while you're at it" that tells it to print out the appropriate ticker name as well.

### 5. Finally! We automatically fit every column to fit the text, so it looks pretty. I used
```
ws.Columns("A:Q").AutoFit
```
I had to place it at the bottom, or it would run too quickly before the numbers had populated. It can make for a funky-looking spreadsheet if it's not at the bottom.

That's it! Thanks for reading. :)
