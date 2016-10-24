# ebay_publisher

### Create a virtual environment and activate:
```
virtualenv env
source path/env/bin/activate
```

### Install python deps:
```
pip install -r requirements.txt
```

### Register on eBay Developer Program:
Go to [ebay developer program](https://developer.ebay.com/join/) and register a new app, then put the config in ```api_config.yaml```.

### To support Google Sheets:
Do the same on [google developers console](https://developer.ebay.com/join/), make a new project and put the config in ```gsheets_credentials.json```.


### Run the script:
```
python ebay_publisher.py -file ./myfile.xlsx
```


### For google sheets:
```
python ebay_publisher.py -url https://docs.google.com/spreadsheets/d/1sdspuxgqym4-9WjwxH87k8ZPGteQgib8uZau5hlRkQM
```
