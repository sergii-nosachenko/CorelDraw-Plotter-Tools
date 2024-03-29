![Project cover](images/cover.jpg?raw=true)

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

## Possible issues and troubleshooting

Sometimes CorelDraw doesn't allow to store workspace or script in system `Program Files` folder due to insufficient user rights.

If you can't install script/workspace or CorelDraw crashes try to manually move script and workspace to Common app folder:

GMS file to
```bash
%AppData%\Corel\{Your CorelDraw version}\Draw\GMS\
```

CDWS file to
```bash
%AppData%\Corel\{Your CorelDraw version}\Draw\Workspace\
```

Restart your CorelDraw and find PlotterTools workspace in `Windows > Workspace` menus.

## Plugins list:

- [Correct table for cut](#correct-table-for-cut)
- [Create table for cut](#create-table-for-cut)
- [Multiply objects for cut](#multiply-objects-for-cut)
- [Prepare curves for cut](#prepare-curves-for-cut)
- [ReDefine frame and reference points](#redefine-frame-and-reference-points)
- [Calculate curves length](#calculate-curves-length)

***

## Correct table for cut

![Correct table for cut](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/icons/icons-01.png?raw=true)

Simple plugin that takes your selected `Table object` or similar set of horizontal and vertical lines and reorganize them in zigzag style for minimize knife idle time and overall cutting time.

**Usage:**

- create table by `Table tool` with desired parameters / or create set of horizontal and vertical lines forming a table / or call [Create table for cut](#create-table-for-cut) and it will create and correct table automatically;
- define frame guide and reference points by [ReDefine frame and reference points](#redefine-frame-and-reference-points) plugin or skip this and plugin will use your documents bounds as ones;
- select your table;
- click `Correct table for cut` icon on `Plotter tools` toolbar.

There no settings for this plugin.

***

## Create table for cut

![Create table for cut](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/icons/icons-02.png?raw=true)

Plugin creates new table with parameters. It will be automatically optimized for cut.

**Usage:**

- select shape for reference as frame *(Note: will be removed after execution)* / or define frame guide and reference points by [ReDefine frame and reference points](#redefine-frame-and-reference-points) plugin or skip this and plugin will use your documents bounds as ones;
- click `Create table for cut` icon on `Plotter tools` toolbar.

**Settings:**

![Settings window](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/images/CreateTableForCut-01.png?raw=true)

- **Cell width**: cell width in mm.
- **Cell height**: cell height in mm.
- **Columns**: expected columns count (calculates automatically to fit frame, can be adjusted to smaller value >= 1).
- **Rows**: expected rows count (calculates automatically to fit frame, can be adjusted to smaller value >= 1).
- **Overcut**: lines offset value in mm (0 - 10).

If cell doesn't fit into frame you will not be able to proceed. Check you frame size and adjust settings.

After operation complete you will get the message with calculated total cut length that can be copied.

**Result:**

![Result table](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/images/CreateTableForCut-02.png?raw=true)

***

## Multiply objects for cut

![Multiply objects for cut](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/icons/icons-03.png?raw=true)

Plugin dublicates selected shape(s) with parameters. It fills frame in zigzag style starting fron bottom right corner.

**Usage:**

- define frame guide and reference points by [ReDefine frame and reference points](#redefine-frame-and-reference-points) plugin or skip this and plugin will use your documents bounds as ones;
- select shape(s) to dublicate;
- click `Multiply objects for cut` icon on `Plotter tools` toolbar.

**Settings:**

![Settings window](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/images/MultiplyObjectsForCut-01.png?raw=true)

- **Horizontal offset**: positive number of offset on x-axis in mm. Recommended to set in to: `cut contour width` + `gap between contours`.
- **Vertical offset**: positive number of offset on y-axis in mm. Recommended to set in to: `cut contour height` + `gap between contours`.
- **Columns**: expected columns count (calculates automatically to fit frame, can be adjusted to both smaller and bigger values).
- **Rows**: expected rows count (calculates automatically to fit frame, can be adjusted to both smaller and bigger values).
- **Align to frame center** *(new)*: align generated contours group to frame center, otherwise leave at same position.

Result bounds can be larger than frame, keep it in mind.

**Result:**

![Result table](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/images/MultiplyObjectsForCut-02.png?raw=true)

***

## Prepare curves for cut

![Prepare curves for cut](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/icons/icons-04.png?raw=true)

Plugin optimize selected shape(s) with parameters. It process all selected shapes one by one, adding smoothness, reducing nodes and rounding corners.

**Usage:**

- select shape(s) to optimize;
- click `Prepare curves for cut` icon on `Plotter tools` toolbar.

**Settings:**

![Settings window](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/images/PrepareCurvesForCut-01.png?raw=true)

- **Curve smoothness level**: can be one of:
    - *None* - no smoothing at all
    - *Low* - minimum level (by default)
    - *Medium* - average level
    - *High* - maximum level

    These parameters have been tuned experimentally based on practice.

- **Maximum corners radius**: can be one of:
    - *None* - no corners fillet radius at all
    - *0,5 mm* (by default)
    - *0,75 mm*
    - *1 mm*

    These parameters have been chosen based on practice.

- **Number of iterations**: can be one of:
    - *1* (by default)
    - *2*
    - *3*

- **Advanced optimization**: if switched ON (by default) works with every segment of curve. Can achive better results but takes longer. Also can cause artifacts in some cases, so it recommended to inspect all curves after processing.
Also this parameter will automatically close all curves.

Feel free to experiment with parameters to achive best results.

**Result:**

You can see the difference between original (magenta) and optimized (green) curves.

![Result table](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/images/PrepareCurvesForCut-02.png?raw=true)

***

## ReDefine frame and reference points

![ReDefine frame and reference points](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/icons/icons-06.png?raw=true)

Simple plugin that takes your selected shape and creates frame guide with reference points. Useful when you have template file with cutting marks for your cutter device. Just create rectangular shape inside available cutting area and plugin adds necessary guides.

Original shape will be removed after completed. Previous frame and reference points will be replaced.

**Usage:**

- create rectangular shape inside available cutting area;
- click `ReDefine frame and reference points` icon on `Plotter tools` toolbar.

There no settings for this plugin.

***

## Calculate curves length

![Calculate curves length](https://github.com/sergii-nosachenko/CorelDraw-Plotter-Tools/blob/master/icons/icons-05.png?raw=true)

Simple plugin that takes your selected shape(s) and calculate total curves length. If object can't be converted to a curve it will throw an error.

**Usage:**

- select all shapes you want to measure (avoid images and complex shapes);
- click `Calculate curves length` icon on `Plotter tools` toolbar.

There no settings for this plugin.