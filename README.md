# GOOGLE SHEET TO CSV
Ruby script to download a range of data from a single Google Sheet and output it as a CSV

## Setup
1. Follow [these instructions from Google](https://developers.google.com/sheets/quickstart/ruby) to get the necessary client_secret.json and gem installation instructions. You may need to rename the file from Google.
2. Change the spreadsheet ID in [quickstart.rb:L50](https://github.com/tongrhj/google-sheet-to-csv/blob/master/quickstart.rb#L50)
3. Change the [data range](https://github.com/tongrhj/google-sheet-to-csv/blob/master/quickstart.rb#L51) and [column numbers](https://github.com/tongrhj/google-sheet-to-csv/blob/master/quickstart.rb#L55) to what you want to extract

## Run It
```
ruby quickstart.rb
```

## Errors?
You may need to [Install Ruby](http://installrails.com).
