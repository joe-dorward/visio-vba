# Shaper 
This is a proof-of-concept, experimental project, to use a Visio VBA application to iterate over a DITA XML file to add shapes and connectors to a Visio drawing.

![shaper_v4_(dita).png](shaper_v4_(dita).png)

## The Process
```mermaid
  flowchart LR

    %% Blocks

    DITA["DITA XML - Topic<br/>(file)"]
    style DITA fill:cornsilk,color:dodgerblue

    VBA["VBA Application"]

    VISIO["Visio<br/>(drawing)"]
    style VISIO fill:cornsilk,color:dodgerblue

    %% CONNECTIONS
    
    DITA-- iterated by -->VBA
    VBA-- to generate -->VISIO

```

**Note** - the project files below 'require' that you have (a) Visio installed, and (b) the knowledge (or tenacity to figure out how) to import, and run VBA code from within Visio.

## Files
* ```t_shapes.dita``` - is the (DITA XML - Topic) with a table containing the information about the shapes to be drawn
* ```shaper_v4_04.bas``` - this is the VBA code module that will iterate over ```t_shapes.dita```, and add the shapes to an open Visio drawing

## Steps

**Step 01** - Download ```t_shapes.dita```

**Step 02** - Download ```shaper_v4_04.bas```

**Step 03** - Import ```shaper_v4_04.bas```, and run the ```Main()``` method

## Visio Drawing
![shaper_v4_(vsd).png](shaper_v4_(vsd).png)
