# Warframe Reward selecter

This program is designed to help you select the reward you want from the rewards selection after completing a void mission. It will read the text from the screen using Optical Character Recognition, and attempt to pull up information about each item from that. The program takes a screenshot to read data so that no code is injected, or read from RAM. I know this is not the most optimal solution, but I did not want to violate any of warframes terms of service.

I developed this program because I kept forgetting the platinum and ducat values of each item when I was at the rewards screen, and wether or not I had as many of that item as I wanted. This program can be used to keep track of the items you want, and how many you have. All information is saved to a file whenever you adjust any values. 

It will display the following values. Platinum, ducats, number of items you own, if you want the item, if the item is vaulted.

How do I run the program?

When you are at the end of the mission, launch the program and click the "run" button while you are at the rewards screen. 
Make sure nothing is blocking the names of all the items when you click the button.

How do I keep track of items that I want?

To add or remove an item from the wantlist, press the "search item" button and type in the name of the item you want to view.
From here, you can press the "wanted" button to toggle the status of whether to track the item, or not. You can also do this from the main screen when the program reads items from the screen

You can press "view list" to view all of the items you have added to your wantlist.

Features still in progress:

- changing the number owned of an item
- adding a scroll bar to the "view list" window
- add functionality to "add to wantlist", currently search item is how items are added
- items with names that span more than one line are sometimes not detected
- make the interface prettier and more intuitive
- add functionality for updating the prices of items as the market changes
