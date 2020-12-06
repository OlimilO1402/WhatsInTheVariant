# WhatsInTheVariant
## Answering the question what is in my Variant
--- Project start 2019-jun-22 --- 
A small project trying to give an answer to the ever so much burning question 
"what the heck is in my Variant?". See Function MVT.VarType2_ToStr(var)
With a testing environment for this function, and along the way trying to 
reveal the question completely.
If you work with COM-Interfaces from Microsoft in VBC chances are high that 
you will come across a Variant-type that is not natively supported by VBC. 
Now you are able to tell which datatype you get from a specific Interface-
function.

--- Update 2019-jul-13 ---
Now with enhanced Enum EVbVarType and converting Variant to and from all other 
datatypes. In VB non intrinsic Variant-datatypes like unsigned types you can 
convert to and from, you can move, copy, and pass to functions but you can not 
us it for calculations.

--- Update 2019-aug-06 ---
This wouldn't be complete without showing the use of IRecordInfo. If Your udt is 
public from an axdll oder typelib, VB can handle it in Variants, where automati-
cally a IRecordeInfo-object is created along with the pointer to the heap-allo-
cated udt-variable, simply by copying (assigning) a udt-variable to a Variant.
