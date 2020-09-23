<div align="center">

## clsStackCollection


</div>

### Description

Cls StackCollection

Use VB's Collections to build stacks and queues. Also included a safety wrapper for collections.

To operate as a stack/queue you do not need a separate key, everything is done using the index

This class adds new Methods and Properties to VB collections.

ArrayFromCollection  	Put contents of stack into an array

ArrayToCollection   	Load members of an array into the stack

Bottom:    		Returns first value added to stack

Clear:     		Empty stack in one call

Exists:    		Tests if index/Key is in stack. Uses error trap to test Optionally runs silent if error occurs

Extract:    		Remove any member from stack

Invert:    		Reverse order of stack

Max:      		Returns the highest alphanumeric member of stack

Middle:			Returns Middle member of stack

Min:      		Returns the lowest alphanumeric member of stack.

Pop:      		Use collection as a LIFO (Last In First Out) stack, removes member from stack.

Pull:     		Use collection as a FIFO (First In First Out) stack, removes member from stack.

Push:     		Wrapper for VB standard Add. Push is used by both LIFO & FIFO stacks.

QuickSort:   		Sort stack into alphanumeric order

RandomMember: 		Extract a random member of the stack(Optional RemoveIt as Boolean = False); if True remove from stack

Replace:    		Change the Item for an Index/Key

Shuffle:    		Randomise the contents of the stack

Top:      		Returns last value added to array

I have also included a class ClsSafeCollection which wraps the standard collection in a safety net

Add:  		Modified to allow you to cope with the various error conditions that can occur

Count: 		VB standard Collection.Count

Item:  		Modified to allow you to cope with the 'Index/Key does not exist' error

Remove: 		Modified to allow you to cope with the 'Index/Key does not exist' error

Please comment and vote.

May just have reinvented the wheel for many of you but hope you find some use for it.

Feel free to use all or parts of this but leave copyrights in place and let me know about it at

rojagilkrist@hotmail.com
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-09-11 12:15:04
**By**             |[Roger Gilchrist](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/roger-gilchrist.md)
**Level**          |Advanced
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[clsStackCo1299719112002\.zip](https://github.com/Planet-Source-Code/roger-gilchrist-clsstackcollection__1-38896/archive/master.zip)








