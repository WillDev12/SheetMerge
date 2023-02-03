# SheetMerge

SheetMerge allows you to program multiple Google Sheets without the hassle of defining hundreds of things.

# How to install

Here is a brief tutorial on how to install. If you are interested in how to use, see [here]()

1. Go to Google Sheets

2. Open new Google Script through Extensions

3. Once in the script, click the plus button above "Libraries"

4. Paste the following ID and press look up:

```
17-iwEaZVMqeY7-7jnTyI9lu4ThNP8fpPzf4AEb93IdRkJKvIcJ6S4SO9
```
5. Select a version and apply

# How to use

SheetMerge provides custom commands that let you define what you want to do to another Google Sheet.  In other words, you can edit two at once.

First, you will need a link.  In my case, I will use %link% as a replacement for an actual one.

Here is an example of how to pair a new sheet:

``` javascript
var link = "%link%";

function yourFunction() {

  var active = SheetMerge.getActiveSheet(link);
  Logger.log(active);
  
}
```

If the logger logs the active sheet used in the link, than the program is fully installed.

There are many more things you can do besides this, and [here]() is a proper list of all of them.

- - -

Here is a script on how to get the properties of a cell:

``` javascript
var link = "%link%";

function getValues() {

  var text = SheetMerge.getValue("A1", link),
      font = SheetMerge.getFontFamily("A1", link),
      size = SheetMerge.getFontSize("A1", link);

  Logger.log(text + ", " + font + ", " + size);

}
```
