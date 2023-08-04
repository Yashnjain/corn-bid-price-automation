1)Setup python and packages on the machine (it might be that stuff is already there):
#https://www.python.org/downloads/
install python-3.8.5-amd64.exe
Install the required packages using pip install -r requirements.txt

2)In the config file check if correct file_locations are present at API_KEY column or not.

3)Clone Clean code from https://gitlab.com/bu_it/cron-bid-price-automation/-/tree/dev?ref_type=heads

4)Navigate to the cloned repository directory and run corn_bid_price_scraper_alert.py script to start the application.

*** This project contains the automation script in python to update the cron bid price sheet for 138 website. ***