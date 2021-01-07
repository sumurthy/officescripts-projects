# Miscelaneos tips

## Full screen editing

Office Scripts doesn't yet offer a full-screen detachable editor. In the meantime, you can use the following steps to use the VS Code client editor to get code completion. It's just for editing purpose only - can't execute script using this. In order to execute, you'd still have to copy/paste code in the Office Scripts in-built code editor (bummer). 

1. In case you already don't have it, [download and install VS Code editor](https://code.visualstudio.com/)
1. Download the intellisence (code completion assistance) [file](index.d.ts) to your local machine into a base folder from where you'll edit your script files. 
1. You can see below that I have saved this file to a folder called OS. 
1. Edit or copy your script to a new file inside of the same folder. The key is to name the script file locally with .ts extension. Otherwise, the code completion doesn’t work. This only works with VS Code editor. 
1. Next add triple-slash directive to the script on line-1. `/// <reference path="{path to index.d.ts}" />`. Example: `/// <reference path="index.d.ts" />`
1. Then, I can start authoring the script and I get code completion. Change the `path to index.d.ts` depending on where your script is. See the image below where the `index.d.ts` is one level above and the `path` directs to it. 
1. Of course, I can’t test or run the code from here. For that I still have to copy/paste the code back to Code editor within Excel on the browser. 
1. Make sure you remove the `/// <reference path="{path to index.d.ts}" />` line before copying code back to Office Scripts code editor in Excel.
1. This is a not so great workaround.. but meets the need to just edit in a larger surface area. I personally use it when I’m editing a large script. 

![Editor](editor2.png)



