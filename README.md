# Alma_append_Fulfillment_Note

Script to append text to Fulfillment Notes in Alma item records


## Requirements
Created and tested with Python 3.6; see ```environment.yml``` for complete requirements.

Requires an Alma Bibs API key with Read/write permissions. A config file (local_settings.ini) with this key should be located in the same directory as the script and input file. Example of local_settings.ini:

```
[Alma Bibs R/W]
key:apikey
```


## Usage
```python append_fulfillment_note.py input.xlsx```
where ```input.xlsx``` is a spreadsheet listing Item IDs in the second column, Holdings IDs in the third column, and MMS IDs in the fourth column.


## Contact
Rebecca B. French - <https://github.com/frenchrb>
