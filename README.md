# Head and Neck Dose Index Picker
This is the script written in C# which is intended to use in Eclipse(Varian, Inc.)<br>
This works like,<br><br>
1. Pick plan, structure set and dose in scope<br>
2. Calculate dose the indices (maximum dose or mean dose or D**% or etc.)<br>
3. Calculate Lyman-Kutcher-Burman's NTCP using DVH and trapezoidal integration.<br>
4. Visualize these indices and judge (O or X) according to the criteria which is determined by the user.<br>
