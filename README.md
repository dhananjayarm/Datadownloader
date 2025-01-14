This program allows you to download data of the symbols given in config.json

Explanation of the fields in config.json

**API_KEY :** API Key as given by Polygon

**start_date :** start date from which data needs to be downloaded.

**end_date :** end date till which data needs to be downloaded. 

**output_dir: **Directory in which file needs to be created

**output_file :** Excel File name where the data needs to be downloaded. If the file is already there, it is append the additional data after the last record currently present

**timeframe:** Timespan for which data needs to downloaded. Possible values -minute, hour,day,second,week,month

**multiplier:** If you want to 5 minute data, the multiplier need to 5, otherwise 1 for 1 minute
