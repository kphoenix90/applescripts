-- by Kelsey L / github: kphoenix90 / 02.27.2018 update
-- 02.27.2018 - In Excel 16.10, "count large" attribute has been removed. Logic has been rewritten accordingly
--**********************************************************************************************************************
-- IMPORTING EXCEL DATA INTO AN OMNIGRAFFLE TABLE
-- 1) Open both the Excel and Omnigraffle files.
-- 2) In the Excel file, select all of the cells that you would like to copy into an Omnigraffle table.
-- 3) Run this script by clicking the "Run" button above.
-- 4) Answer the dialog that appears to specifiy whether you would like to use row striping.
-- 5) A table should be drawn on the newly created layer "excel data grid".
-- 6) Now you may need to resize certain columns to fit the data accordingly.

--**********************************************************************************************************************

global xlsCols --number of selected columns in the excel doc
global xlsRows --number of selected rows in the excel doc
global numCells
global cellWidth
global cellHeight
global xlsdata

--you may customize these parameters to change the visual style of the Omnigraffle grid
set cellWidth to 120 --size of cell width
set cellHeight to 22 --size of cell height
set cellFont to "Arial" --cell font
set lineColor to {53466, 53465, 53466} -- color of the lines in the table
set cellColor1 to {1, 1, 1} --- default background color of the table
set textColor to {0, 0, 0} --default text color of the table 
set cellColor2 to {57311, 57311, 57311} -- background color of every other row if there is row striping
---end of paramaters that can be changed


--copy data from excel
tell application "Microsoft Excel"
  	set xlsdata to {} --holder of excel data
	set props to properties of selection
	
	---check if excel selection exists
	if selection is null or formula of selection is "" then
		display dialog "You have not made a selection." & return & "Please select the cells you would like to import into Omnigraffle and try again."
		return
	end if
	
	set xlsRows to length of formula of props
	set xlsCols to length of item 1 in formula of props
	set numCells to xlsRows * xlsCols
	
	
	set celldata to ""
	
	--copy data from all selected cells and store in xlsdata
	repeat with j from 1 to numCells
		set celldata to string value of cell j of selection
		
		set xlsdata to xlsdata & celldata
	end repeat
	
	log "end of excel copy"
end tell


--prompt user as to whether they would like row striping
display dialog "Would you like to use row striping?" & return & "('y' or 'n')" default answer ""
set striping_yn to text returned of the result as string
if striping_yn is not equal to "y" and striping_yn is not equal to "Y" then
	set striping_yn to "n"
end if

tell front window of application "OmniGraffle Professional 5"
	set cellX to 0
	set cellY to 0
	set cellNum to 1
	set cellColor to cellColor1
	set selectedCells to {}
	tell canvas 1
		
		--create new layer for grid
		set theLayer to (make new layer at beginning of layers with properties {name:"excel data grid"})
		
		--create cells of grid
		repeat with k from 1 to xlsRows
			repeat with i from 1 to xlsCols
				--draw a single cell
				make new shape at beginning of graphics of theLayer with properties {origin:{cellX, cellY}, size:{cellWidth, cellHeight}, draws shadow:false, name:"Rectangle", text:{size:14, text:item cellNum of xlsdata, font:cellFont, color:textColor}, stroke color:lineColor, thickness:1, fill color:cellColor}
				set cellX to cellX + cellWidth
				set cellNum to cellNum + 1
			end repeat
			
			--if row striping is specified, change cell color
			if striping_yn is not equal to "n" then
				if cellColor is equal to cellColor1 then
					set cellColor to cellColor2
				else
					set cellColor to cellColor1
				end if
			end if
			
			--set the cell coordinates
			set cellX to 0
			set cellY to cellY + cellHeight
		end repeat
	end tell
	
	--select all of the created cells
	set selection to graphics of theLayer
	
	--create table of selection
	tell application "OmniGraffle Professional 5" to activate
	tell application "System Events"
		tell process "OmniGraffle Professional"
			click menu item "Make Table" of menu "Arrange" of menu bar 1
			--click menu item "Make Table" of menu 1 of menu bar item "Arrange" of menu bar 1 ---may need to try this line instead with more recent omni's 
		end tell
	end tell
	
end tell

