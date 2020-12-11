# Jar-ShoppingAssistant-P2
## Project Description
Our automation project is centered around monitoring the market for a given user's stock(s). Wouldn't it be nice to check the performance of your stock(s) at only one point in the day? This can easily be achieved with our Stock Manager automation robot. Moreover, you can spend little to no time obtaining a graphical representation of your stock(s') performance by receiving a daily email from our bot.

## Technologies Used
- UiPath Studio
- UiPath Orchestrator
- Excel
- Outlook
- VB

## Features
- User is able to use our robot is able to interact with the web and monitor a given stock every 30 minutes.
- User obtains an overall day performance of their stock(s) when the market closes via email.
- User will find all information about each stock represented graphically through Excel.

## Getting Started
To Clone: `git clone https://github.com/201019-UiPath/Jar-P2.git`

- You'll need to create a folder within you C:\ drive called **JAR-StockManager**.
- Within this new folder, create two new excel files. One named **Tickers.xlsx**, the other named **StockData.xlsx**.
- Open **Tickers.xlsc** and create two headers called **ID** and **Ticker** within cells **A1** and **B1** respectively. 
- Within the **ID** column starting at the value **1**, write how many Tickers you will be tracking. For instance, if you'll be tracking 20 Tickers, write down the values 1-20 in each cell, ending at cell A21.
- Within the **Ticker** column, write each individual Ticker of the stocks you'll be tracking up to the cell which will be adjecent to your last **ID** number. A Ticker is the abbreviated name of a Company. For instance, when searching Apple Stock, you'll see that the Ticker associated is AAPL.
- Save and close **Tickers.xlsc**. Open **StockData.xlsx**.

## Usage
- To run this project from Orchestrator, follow these steps:
    - Navigate to the **Automations** tab within the **Default** folder.
    - Click **Triggers**, enable the **Dispatcher** and **StockBot Performer** schedules in that order.
- Wait until 5:00 PM EST when stocks close, then check your email.

## Contributors
<a href='https://github.com/rachelhufnagel'>Rachel Hufnagel</a>

<a href='https://github.com/antonyt96'>Anthony Tejada</a>

<a href='https://github.com/jamesPan3'>James Panebianco</a>

## License
This project uses the following license: [MIT License](LICENSE).