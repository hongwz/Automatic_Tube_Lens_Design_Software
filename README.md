# Automatic_Tube_Lens_Design_Software
This software is used to automatically design microscope tube lenses with a pair of stock achromatic doublets.

Software packages were produced with both Matlab 2016b and 2019a.

Doublet Selector "with CARD" can only be used with Zemax version after 20.2, while versions "without CARD" do not have the version requirement.

# Catalogue Generator
This software does not have Zemax version requirement, but the doublet catalogue produced by this software depends on different Zemax versions. 

Since different Zemax versions may provide different lens libraries, the catalogue used in Doublet Selector is supposed to be produced by the same Zemax version.

Input a focal length range, a lens diameter range and select lens vendors, then the Catalogue Generator will produce a doublet catalogue from the OpticStudio lens database.

This catalogue can be directly used in the Doublet Selector.

# Doublet Selector
Doublet Selector 1.0 is produced to be used with Zemax version after 20.2 with CARD operand.

Doublet Selector 1.1 is produced to be used with Zemax version before 20.2 without CARD operand.
