F3J Scoring Program

Release 4.00
32bit Version VB5

Disclaimer & Copyright 
 
This program is Copyright (C) 1996 by Stephen Gloor, 
All Rights Reserved. 
 
They are provided as shareware with the following limitations: 
 
This program is alpha-code and is not to be resold or distributed for sale with other programs which are for sale.  There is no warranty or claim of fitness or reliability. The programs are distributed AS IS, and  the author,  shall not be held liable for any loss of data, down time, loss of revenue or any other direct or indirect damage or claims caused by this program. 

The program has been ported to 32 Bit.  It has not been changed or modified.


Brief Operating Instructions

The program is based around a contest.  The data file is called DATA.MDB and is selected by default. 

The program now allows teams.  It also scores Australian 8-11 Thermal Duration contests

The initial screen has a menu system and a display.  The menu is a follows

Menu Option FILE

Open	 This is the menu item to open a data file other that the default

Print Setup	This brings up the standard windows prompt dialogue box

Print	Brings up the Print Selection Screen.  Using the radio buttons select the type of report.  Then push the PRINT REPORT button.  Only the first 2 are active.

Exit	Either this or the exit button on the bottom of the screen will exit the program.  All data is saved when it is entered


Menu Option	COMPETITION

New Competition	Brings up the form to enter a new competition into the database.  This is the first step in creating a new contest.  Supply a name and start/finish date.  The contest type is selected with the radio buttons. Enter the number of rounds that this contest will have(default is 5) When the contest is entered correctly it is automatically the current contest.

Open Competition	Brings up the selection screen for competitions that exist in the database.  Using the NEXT button select the competition that you would like to work on and then push the SELECT button.  This makes this contest the current contest.


Enter Competitors	Selects the Competitor Entry screen.  In the left box is the list of pilots in the database.  Simply select and existing pilot and push the ADD button.  The pilot will appear in the contest box.  If the pilot is not in the Contest then push the ADD PILOT TO DATABASE.  A Form will appear so that a new pilot's details can be entered.  Once the details have been entered then push DONE and the Pilot in Database Box will have the new pilots details displayed.  You can enter pilots up until you set up rounds after this you have to reset the contest to enter more pilots.(NOTE1) The frequency fields are the frequency in Mhz.  Any other format will be rejected.

Set Up Teams	This displays a screen to set up teams in the competition.  Simply select the pilot in the comp from the left and then push ADD.  The pilot will then appear in the right box.  The team number is selected by pushing the scroll buttons next to the number.  Changing this will change the team selected   If you want to remove a pilot from a team then select the correct team and then the correct pilot.  Push the REMOVE button.  The pilot will disappear from the right screen and re-appear in the left screen which is the pool of unallocated pilots.  A tip here is to ensure that team members are on the same frequency so that the round/slot allocation is easier.
	

Manage Freqs		Select this before you select round setup.  It gives you an idea of the number of people on the frequencies used in the contest.  It allows you to see if there is a particular frequency that has too many pilots.  If there is a freq that is heavily used then it can be selected to move pilots off onto their alternate frequencies.	Select the problem frequency push the button marked "Select pilots on this Freq" and all pilots on this frequency in this contest are displayed.  Select a pilot and push change frequency.  This will cause the freq used for this pilot to change to the alternate.  If a new frequency is entered here it will change the pilots database entry.

Set Up Rounds	Sets up rounds for the current contest. This menu Option will show the Set Up Rounds screen. If the contest already has the rounds set up then the rounds cannot be set up again until the contest has been reset (see TOOLS - Reset Current Contest) .   The initial screen is a blank display screen with 4 Command buttons .  

Set Up Rounds Form - Allocate Slots	This button starts the allocate process.  The progress of the allocation is displayed by the scroll bar along the top.  The resulting round/slots are displayed in the display screen. Check here to see if the number of pilots in each slot is even and there are sufficient  pilots in each slot.   If the draw looks even then accept the draw by pressing the "Accept Allocation and Set up Scores in Database" and all the score entries will be created and this contest will be set up.   If a screen comes up that says "Cannot do allocation ....." just press the set up rounds button again.  It sometimes takes 2 or 3 goes to get a successful allocation.  If it cannot allocate after 2-3 goes you will have to modify the team allocation or shift some pilots of a particular frequency.

Set Up Rounds Form - Change Allocation	If the allocation the computer comes up with is not quite right yet not a total loss you can modify the draw manually.  Pressing this button will bring up the Change Allocations Form

Change Allocation Form	- This form allows you to chop and change manually the computer draw.  Sometimes the computer will make a slot with only 2 pilots in it or something.  Down the left hand side is a box that shows the number of pilots in each slot for each round.  Using the Inc and Dec arrows beside each round and slot number you can select the round and slot with too few pilots.  By pressing the button "Remove Pilot From Slot" you can take the pilot out of the slot.  Then by using the Inc and Dec arrow by the slot number (Notice that the Round arrow is greyed out) you can select the slot that you want this pilot to go into.  The computer will check to see if there is a freq conflict when you push the button "Add Pilot To Slot" and will not let you allocate the pilot if there is.  If there is a Team conflict the you will be given the option to allocate the pilot.  If you remove all pilots from the last slot in every round the computer will remove this slot from all rounds.  This only applies to the last slot. 

	

Set Up Rounds Form - Don't Accept Allocation  This button clears the display and reset the matrix.  Up until you set up the contest by pushing the "Accept Allocation .. " button the draw resides in memory in a matrix and can be changed at any time.

Set Up Rounds Form - Accept Allocation ......	Pressing this button as previously described commits the contest draw to the database.  After this point the allocations can only be changed by resetting the contest.  However a pilot can be moved from slot to slot by selecting "Move Pilot" in the Competition menu

Move Pilot	If a pilot has to be moved during a contest for any reason then select Move Pilot.  The format is exactly the same as in the allocation process however now the database is being modified.  Use with caution as you can change the results of slots and rounds by moving a pilot with scores. 

Leader Board	As the contest progresses you can display the current state of the contest at any time.  It will display in total score order the current results.

Menu Option   SCORES

Enter Slot Flight Times	This brings up a form to enter the flight times.  Select the round and slot that you wish to score and the pilots in this slot will be displayed.  When the flight details for a pilot have been entered push ENTER SCORES FOR PILOT button and the score will be recorded in the database.  The form will then automatically advance to the next pilot. You can navigate between pilots using the NEXT and PREV keys.  If you make a mistake simply navigate to the error and re-enter.  This will overwrite the previous entry. The penalties are calculated automatically from the entered flight time.  ie if you enter 10 mins 30sec then a penalty of 60 points is taken from the score.  It is 60 because the penalty is 30 points and also points stop being recorded at 10:00.  This "penalty" of 60 points will give the correct raw score.  If however a pilot has a relaunch,makes 7mins 40sec and overflies the slot by 29 sec.  If you just enter 7:40 then normally no penalty will be added.  In these cases the penalty will have to be calculated manually and entered in the additional penalties field.  If a entry is present in this field then the program will add the entry to the total penalties. The landing score is auto calculated with the distance from the spot being entered.  After the last pilot's score has been entered push the SCORE SLOT button.  This will calculate the normalised score for the pilots in the slot.  This can be done at any time and you can leave it for later by pressing the NO CALCULATE button.  You can rescore any slot at any time simply be entering the new flight times and then re-scoring.

View Scores	This is a table that shows the scores for the selected round and slot.  The round and slot can be selected with the arrow inc and dec buttons beside the numbers.

Menu Option   DATABASE

Maintenance	This menu option displays a form that allows you to maintain the database.

DataBase Maintenance Form	This form can be selected with no contest selected.  On the right hand side is the COMPETITORS MAINTENANCE SCREEN.  On the RHS is the CONTEST MAINTENANCE SCREEN.

Competitors Maintenance Screen		Selecting a displayed pilot will allow you to 1.  Change the pilots details  2.  DELETE Pilot - Deleting a pilot moves the pilot to a non-current list.  This is important because if you delete a pilot completely from the database you could affect the outcome of a previously stored contest.  This option simply removes the pilot from the current list so that it is not cluttered up with pilots who no longer fly.  3. ADD NEW PILOT  - This option allows you to pre-enter pilots into the database without having a contest selected so that they can be entered into a contest at a later date.  4.PURGE A PILOT FROM DATABASE - If you are sure that the pilot does not need to be in the database then you can PURGE the database.  This removes all references in all contests for this pilot.  You lose all scores in all contests for this pilot.   The Radio Buttons under the screen allow you to change the screen to show  current pilots or pilots that have been 'deleted'.

Contest Maintenance Screen	This screen allows you to maintain contests.  The current contest is denoted by **.  You can change the contest parameters and delete a contest using this screen.

Search Database	This menu option shows a search screen that allow you to search the database for one of the displayed parameters or a whole name.

Menu Option TOOLS

Reset Contest	This enables you to redraw a contest.  If a round/slot allocation is done for a contest then it can not be re-done until it is reset.  Selecting this option will erase any scores for the contest an reset the contest so that it can the menu selection SET UP ROUNDS will work again.

Set Current Contest Complete	Sets the contest complete which does 3 things.  1.  Adds up the final scores for the pilots, dropping the lowest score if more that 5 rounds are flown.    2.  Puts the place the pilot acheived in the database  3.  Locks the contest so that no more scores can be put in.    Note: If for any reason the contest does not run the full course of the number of rounds entered when the contest was entered then you can set the contest complete at any stage by selecting this menu option.

HELP	No help has been done yet except this.



