Attribute VB_Name = "READ_THIS"
'Suggestions for texting the features in this sample project
'
'Add Items:
' - Enter some text in item 1
' - Press the 'Add' Button or hit 'Enter' (cmdAdd.Default is True, so hitting enter
'   triggers this button unless another button has focus
' - The item is added and the text in item 1 is selected
' - Type something else.  The selected text is deleted first automatically.
'
'Change:
' - Type Blue in Item 1 and Azure in Item 2
' - Press The Change Button - 'Blue' is changed to 'Azure'.
' - Press the button again.  Notice it is not found, even though 'Navy Blue'
'   appears on the list.  The whole string must match
'
'Find:
' - Type GREEN
' - Press 'Find Item 1 Exactly'
' - 'Green is in position 5'.  The search is not case-sensetive
' - Green is selected in the list by the code
' - Type Purple
' - Press 'Find Item 1 Exactly'
' - Purple does not appear on the list (Dark Purple does; but the whole string must match)
'
'Find "Begins With Item 1"
' - Type Dark
' - Press 'Find "Begins With.."
' - Dark Red is in item 1.  The button caption changes to 'Find Next "Begins With..."'
'   (the item is not selected because the code was deliberately left out of this routine)
' - Press 'Find Next...'
' - Dark Green is in position 6
' - Press 'Find Next...'
' - Dark Purplr is in position 11
' - Press 'Find Next...'
' - The entire list has been searched. The button is unavailable until you change item 1
'
'Remove
' - Click an item to select it.  The Remove button becomes available
' - Click Remove.  The item is removed and nothing is selected, so 'Remove' is
'   no longer available
'
'ToolTip on hover
' - Hold mouse over 'That Sickly Green-Yellow Color'
' - Item is too big to see, so the name appears in a tool tip
' - Hold mouse over 'That Reddish-Orange Color'
' - Item fits in window.  No tool-tip
' - Add items with the 'Add Item' button until a scrollbar appears in the list
' - Hold mouse over 'That Reddish-Orange Color'
' - Item is now too big to see, name appears in a tool tip
' - Hold mouse over smaller item
' - Tool tip disappears
' - Remove items till scrollbar disappears.  No more tool tip for
'   'That Reddish-Orange Color'
'
'Clear
' - Press Clear
' - Entire list is cleared and 'Clear' button is unavailable (if an item was
'   selected, 'Remove' becomes unavaiable also.
' - Add an Item
' - Clear button becomes available again
