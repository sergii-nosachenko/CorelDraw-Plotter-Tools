# CoreDraw tools for preparing files for plotter cutting
This tiny plugins package will be useful for those, who prepare files for kiss-cut and die-cut postpress operations.
Stickers, shaped cards, boxes etc. But in first place - it's about sheet printing.
Main features are:
- creating and optimizing table-view layouts (rectangular stickers without gap);
- zig-zag style imposition of shapes for minimizing plotter movements;
- optimizing complex shapes for best knife performance.

## Requirements:

- CorelDraw >= 2020

## Installation:

1. Download files from `dist` folder.
1. Open your CorelDraw application.
1. Open `Scripts` docker panel:
    - from menus: `Window > Dockers > Scripts`;
    - or by keyboard shortcut: `Ctrl + Shift + F11`
1. In scripts list select tab `Visual Basic for Applications`.
1. On top of the docker press `Load...` button.
1. In file open dialog navigate to `PlotterTools.gms` file from `dist` folder. Choose open.
1. Now time for command panel. Open `Windows > Workspace > Import` from top menus.
1. In file open dialog navigate to `PlotterTools.cdws` file from `dist` folder. Choose open.
1. Leave all checkboxes enabled. But select to import into `Current workspace` (recommended).

## Plugins list:

### Correct table for cut

![Correct table for cut](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/icons/icons-01.png?raw=true)

Simple plugin that takes your selected `Table object` or similar set of horizontal and vertical lines and reorganize them in zigzag style for minimize knife idle time and overall cutting time.

**Usage:**

- create table by `Table tool` with desired parameters / or create set of horizontal and vertical lines forming a table / or call [Create table for cut](#create-table-for-cut) and it will create and correct table automatically;
- define frame guide and reference points by `ReDefine frame and reference points` plugin or skip this and plugin will use your documents bounds as ones;
- select your table;
- click `Correct table for cut` icon on `Plotter tools` toolbar.

There no settings for this plugin.

### Create table for cut

![Create table for cut](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/icons/icons-02.png?raw=true)

Plugin creates new table with parameters. It will be automatically optimized for cut.

**Usage:**

- select shape for reference as frame *(Note: will be removed after execution)* / or define frame guide and reference points by `ReDefine frame and reference points` plugin or skip this and plugin will use your documents bounds as ones;
- click `Correct table for cut` icon on `Plotter tools` toolbar.

**Settings:**

*In progress...*