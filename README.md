# outlook-to-telegram
Script that forwards outlook mail to the telegram channel

## Usage

1) Clone this repository:
   
   `git clone https://github.com/egormkn/outlook-to-telegram.git && cd outlook-to-telegram`
   
2) Install dependencies:
   
   `npm install`
   
3) Build javascript files:
   
   `npm run build`
   
4) Run script for the first time to set some preferences:
   
   `npm run start`
   
5) Set the CRON job to start the script every 10 minutes:
   
   `(crontab -l ; echo "*/10 * * * * (cd $(pwd) && npm run start)") | crontab -`
   
