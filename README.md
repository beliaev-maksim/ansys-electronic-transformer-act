## About
The Electronic Transformer Extension is a customized solution for electronic transformers.  The ACT provides an 
easy-to-use interface to draw the geometry and setup a solution for a transformer or inductor.  A database of basic 
topologies and materials for the commonly used cores is included in the ACT which allows users to choose the required 
shape and size of the core. In addition, users can define their own winding strategy using Winding definition panel 
(Planar or Wound types) which enables automatic creation of all winding turns with rectangular or circular 
(only for wound type) cross section.
The ACT allows users to automatically setup an Eddy Current Solution with (or without) a Frequency Sweep Definition. 
The ACT considers the frequency dependent core permeability and core loss Steinmetz coefficients. The ACT also allows 
defining matrix connections (series or parallel) if required. This solution process creates a frequency dependent R/L model which can be imported into 
ANSYS Simplorer as a Maxwell Dynamic Eddy Current component.  

More information you can find on the [help page][1]  
Software is distributed under [GNU License](LICENSE)

 [1]: src/ElectronicTransformer/help/help.html
 
 ## Installation
 ### Supported versions
 Transformer Toolkit v2.0 supports Ansys Electronics Desktop of version 2021R1 and higher
 
 ### Installing ACT
 Note: you can find instruction with pictures in [overview PDF][1]  
 To install the app you need to 
 1. open Electronics Desktop
 2. Navigate to menu _View_ and activate _ACT Extensions_
 3. Go to Manage Extensions and click "+" sign
 4. Select .wbex file that you can downloaded from GitHub Releases page
 5. new ACT will appear in the menu, click on it to activate
 6.  Now you can go oage back and go to Launch Wizard panel
 
 [1]: doc/ETK_2021R1.pdf
 
 ## Examples
 Tranformer Toolkit ACT is provided with basic examples based on public papers and articles. You can open examples from 
 ACT itself by clicking _Open Examples_ button
 
 ## Contribution
 Current projects welcomes new contributors. You can contribute in one of the following ways:
 1. Provide additional examples
 2. Contribute to the code by fixing some bugs or implementing new features
 3. Provide additional default core dimensions or material parameters
 4. Adding new tests that are run against Ansys Official and Pre Releases
 
 Note: current software is distributed under [GNU LICENSE](LICENSE)
 
 ### How to contribute
 1. Make a fork of current repository
 2. Clone your fork to local machine
 3. Implement changes in new branch 
 4. Push your branch to your fork
 5. Create a Pull Request for merging
 
 Note: all code changes should follow as much as possible PEP8 guideline and all unittests should pass validation
 
 ### Bugs and Issues
 If you find some bug or wrong data please open an issue on GitHub project page