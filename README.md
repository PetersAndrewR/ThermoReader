# ThermoReader

ThermoReader allows for the recording of temperatures and experiment times from a network connected YOKOGAWA DX100p.  The recorded information is then written to a newly/appropriately named excel file where the data can be further analysed between experiments to identify trends in heating.

**Instructions**
  -Main Window-
The main window contains three buttons on the left.  Starting from the bottom, The "Example Output" button opens an example excel sheet to show what the final output of a test is.  This example was an actual test result during testing of ThermoReader.

The "Open ReadMe" button opens this file.

The "Start Test" button is used to start a new test on one of the probes.  Pressing this button opens a new window which request information about the test and does not start the timer on it's own.

The 6 green check mark images are mapped to the 6 probes and it is in these regions we can see the current status of the probe.  The green check image is the ready indicator used to inform the user the probe does not currently have a test running.  This image can be replaced by a blue circle with 3 smaller white circles inside to indicate a test is currently running on that probe, a red circle with a flame to indicate the experiements Working Time has been met, and a red cirlce with a dash inside indicating a probe error has been detected.  These images are non-clickable and will update on their own.

  -New Test Window-
The New test Window is created after pressing the "Start Test" Button.  In this window are some text boxes and radio buttons.  This window asks for Product Name, Lot#, and the Test # or Letter.  At the bottom, using the radio buttons, you select the probe number the test will be started on.  Only once the information is enetered, and the "Ok" button is pressed does the test and the timer start. Once the ok button is pressed, the New Test window is automatically closed and you will see that the status of the selected probe on the main window is updated to show a test is currently running.

**Good to Know**
Currently, a test is ended once the Current Temperature of the probe is 10 degrees less than the max recorded temperature for the current test.  If a test is ended early, such as when the temperature has only changed a few degrees, the only way to end the test to free up the probe for another test is to either close the program completely or to dip the probe in cold water to force a temperature change of more than -10 degrees.

**Credits**
Icons credited to http://fasticon.com/ and explicitly available at http://www.iconarchive.com/show/action-circles-icons-by-fasticon.html

**Planned Featurs / Still to Come"**
-Test to ensure New test text boxes are filled out and a probe is selected
-Test to ensure a new test isn't started on a currently running probe
-New window creation to see progress of currently running tests such as time and temp
-Ability to end a test early through the program instead of dipping the probe in cold water

Written in Python 3.6.5
