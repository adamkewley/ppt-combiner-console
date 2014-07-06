# PPTCombiner
*Combine multiple presentation files into one.*

# ![Screenshot of PPTCombiner](https://github.com/AdamK117/ppt-combiner-console/raw/master/img/screenshot.png)

A console-based program that automates PowerPoint to merge multiple presentation files. Suitable for regular meetings and presentations where a little automation goes a long way. A more straightforward, GUI driven, approach can be found at the [Presentation Combiner](https://github.com/AdamK117/presentation-combiner) project.

This project is inspired from the excellent [PPTJoin](http://code.google.com/p/powerpointjoin/) (updated perl-based version [here](https://github.com/richardsugg/PowerpointJoin)) and offers some extra functionality. Presentation files and folders containing them can be used as merge targets with `PPTCombiner`. 

**You must have PowerPoint installed for `PPTCombiner` to work.**

## Quick Start

  - Copy `PPTCombiner` into a folder containing the presentation files you want to combine. 
  - Open a console (`SHIFT+Right Click -> Open Command Window Here`) 
  - Type `CScript pptCombiner.js` and run it (`Enter`).
  - The combined presentations will appear as `combined.ppt` in the folder.

## Directory Argument

  - Run `PPTCombiner` with a directory as its argument: `CScript pptCombiner.js C:\path\to\directory\`
  - The combined presentations will appear as `combined.ppt` in your working directory.
  - You can also use relative paths, e.g. `CScript pptCombiner.js some\relative\path`.

## Text Argument

*example.txt*
    
    C:\An\Absolute\Path.ppt
    A_Relative_File.ppt
    \A\Relative\Path.ppt
    C:\An\Absolute\Directory
    A\Relative\Directory

  - Run `PPTCombiner` with the `.txt` as the argument: `CScript pptCombiner.js example.txt`
  - The combined presentations will appear as `combined.ppt` in your working directory.

## Change the Output File

  - Use the `/o:` flag to define an output path e.g. `CScript pptCombiner.js /o:C:\output-here.pptx` C:\some-input-directory\
  - The combined presentations will appear as `output-here.pptx` in the `C:\`.

## Other Flags

  - `/?`: View in-terminal readme.
  - `/R`: Enable recursive mode. `PPTCombiner` will recursively look for any presentation files in any supplied directories.
