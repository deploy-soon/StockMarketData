# StockMarketData
Drive stock market data with CybosPlus

## Preliminaries
> http://cybosplus.github.io/
- Open an account
- Download hts-cybos


## How to use

```
$ pip install -r requirements.txt
$ cd market
$ python timeseries.py --conf config/timeseries.json
    
# get dart data
$ python dart.py run --days=:days run
$ python dart_report.py Danil run
$ python dart_report.py Usang run
$ python dart_report.py Treasury run
$ python dart_report.py CB run

$ python dart_minute.py run
```
