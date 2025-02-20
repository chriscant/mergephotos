# mergephotos

Command line tool to find photo filenames and add into matching species list spreadsheet rows.

This is used for the Flora of Cumbria to add columns to a species list spreadsheet.
Two columns are added for each matching species photo:

* the relative path to the photo filename
* a photo caption - based on the photo filename

usage: `node mergephotos.js <config.json>`

Where the config file looks like this:

```
{
  "input": "data/input.xlsx",
  "photosdir": "photos directory",
  "output": "data/output.xlsx"
}
```

The code automatically fixes potential issues such as times sign × and letter x and photo folder name anomolies.

# Licence

[MIT](LICENCE)
