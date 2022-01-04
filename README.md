# Automatic_Tube_Lens_Design_Software
This software is used to automatically design microscope tube lenses with a pair of stock achromatic doublets.

Software packages were produced with both Matlab 2016b and 2019a.

Doublet Selector "with CARD" can only be used with Zemax version after 20.2, while versions "without CARD" do not have the version requirement.

# Catalogue Generator
This software does not have Zemax version requirement, but the doublet catalogue produced by this software depends on different Zemax versions. 

Since different Zemax versions may provide different lens libraries, the catalogue used in Doublet Selector is supposed to be produced by the same Zemax version.

Input a focal length range, a lens diameter range and select lens vendors, then the Catalogue Generator will produce a doublet catalogue from the OpticStudio lens database.

This catalogue can be directly used in the Doublet Selector.

PRESS "CHECK THE CONNECTION WITH ZEMAX" TO INITALISE THE CONNECTION BETWEEN ZEMAX AND MATLB FIRST.

After entering all the catalogue requirements, press "Generate" and "Refresh" buttons to produce desired doublet catalogue.

# Doublet Selector
Doublet Selector 1.0 is produced to be used with Zemax version after 20.2 with CARD operand.

Doublet Selector 1.1 is produced to be used with Zemax version before 20.2 without CARD operand.

PRESS "CHECK THE CONNECTION WITH ZEMAX" TO INITALISE THE CONNECTION BETWEEN ZEMAX AND MATLB FIRST.

1. Enter the tube lens design parameters: tube lens focal length, entrance pupil diameter and field angle.

2. Set the lens separation range according to the system requirement.

3. Choose the optimisation option in "Distance from Entrance Pupil (EP) to Doublet 1"

(if you use the version without CARD to optimise "Telecentric" or "Set distance from EP to 1st principal plane", then you need to select the optimisation accuracy.)

4. Set the tube lens configuration. According to our research, based on stock optics Configuration C can provide the best performance.

5. Set the optimisation reference (Chief Ray or Centroid)

6. After entering all the requirements, press "Go" and "Print" buttons to start the optimisation and print out the results in the user interface.

"Decode" button is used to decode the lens pair number.

See "TEST SAMPLE" folder for more information.
