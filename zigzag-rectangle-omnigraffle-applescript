--kelsey L / kphoenix90 / 2013


---script to create zig zag square based on the dimensions of a selected shape. Can specify which sides of the square are zig zag.
--max width is 732px
--omnigraffle selection must be made
--user input required for which sides are cropped

global x_val
global y_val
global ctr
global limit
global zag_depth
global zag_length
global orig_y
global shape_origin
global selecteD_shape
global obj_size

tell front window of application "OmniGraffle Professional 5"
  set obj_size to {}
	set shape_origin to {}
	set selecteD_shape to selection
	repeat with obj in selecteD_shape
		set obj_size to size of obj
		set shape_origin to origin of obj
	end repeat
	
	if length of selecteD_shape = 0 then
		display dialog "You have not made a selection. Please select your crop and try again."
		return
	end if
	
	--figure out cropping sides
	display dialog "Are you cropping the top?" & return & "('y' or 'n')" default answer ""
	set top_yn to text returned of the result as string
	if top_yn is not equal to "y" and top_yn is not equal to "Y" then
		set top_yn to "n"
	end if
	
	display dialog "Are you cropping the right?" & return & "('y' or 'n'))" default answer ""
	set right_yn to text returned of the result as string
	if right_yn is not equal to "y" and right_yn is not equal to "Y" then
		set right_yn to "n"
	end if
	
	display dialog "Are you cropping the bottom?" & return & "('y' or 'n'))" default answer ""
	set bottom_yn to text returned of the result as string
	if bottom_yn is not equal to "y" and bottom_yn is not equal to "Y" then
		set bottom_yn to "n"
	end if
	
	display dialog "Are you cropping the left?" & return & "('y' or 'n'))" default answer ""
	set left_yn to text returned of the result as string
	if left_yn is not equal to "y" and left_yn is not equal to "Y" then
		set left_yn to "n"
	end if
	
	--set zig zag dimensions and find limits for polygon size
	log "shape origin " & shape_origin
	log "object size " & obj_size
	
	set x_val to item 1 of shape_origin
	set y_val to item 2 of shape_origin
	set orig_x to x_val
	set orig_y to y_val
	set ctr to 1
	
	set zag_depth to 4
	set zag_length to 14
	
	set shape_width to item 1 of obj_size as integer
	set shape_height to item 2 of obj_size as integer
	repeat while shape_width mod zag_length is not 0 and ctr < 12
		set shape_width to shape_width + 1
		set ctr to ctr + 1
	end repeat
	--log "shape width " & shape_width
	if shape_width > 732 and right_yn = "n" then
		set shape_width to 732
	end if
	if shape_width > 732 and right_yn is not equal to "n" then
		set shape_width to 713
	end if
	log "shape width " & shape_width
	
	set ctr to 1
	repeat while shape_height mod zag_length is not 0 and ctr < 12
		set shape_height to shape_height + 1
		set ctr to ctr + 1
	end repeat
	set ctr to 1
	
	set x_limit to ((shape_width / zag_length) as integer) + 1
	set y_limit to ((shape_height / zag_length) as integer) + 1
	log "x limit " & x_limit
	log "y limit " & y_limit
	
	
	
	
	
	set point_list to {}
	set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
	--if top is cropped, do zig zag
	if top_yn is equal to "y" or top_yn is equal to "Y" then
		---top
		repeat while ctr < x_limit
			set x_val to (x_val + zag_length)
			if ctr mod 2 = 1 then
				set y_val to (y_val + zag_depth)
			else
				set y_val to (y_val - zag_depth)
			end if
			set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
			set ctr to ctr + 1
		end repeat
		if ctr mod 2 = 1 then --ends up, then return to baseline
			set y_val to y_val + zag_depth
			--set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
		end if
	else --if side shouldn't be cropped, starting and end points
		--set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
		set x_val to x_val + ((x_limit - 1) * zag_length)
		set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
	end if
	
	
	--right
	
	set ctr to 1
	if right_yn is equal to "y" or right_yn is equal to "Y" then
		repeat while ctr < y_limit
			set y_val to (y_val + zag_length)
			if ctr mod 2 = 1 then
				set x_val to (x_val + zag_depth)
			else
				set x_val to (x_val - zag_depth)
			end if
			set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
			set ctr to ctr + 1
		end repeat
		if ctr mod 2 = 0 then
			set x_val to x_val - zag_depth
			--set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
		end if
	else
		set y_val to y_val + ((y_limit - 1) * zag_length)
		set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
	end if
	
	---bottom
	set ctr to 1
	if bottom_yn is equal to "y" or bottom_yn is equal to "Y" then
		repeat while ctr < x_limit
			set x_val to (x_val - zag_length)
			if ctr mod 2 = 1 then
				set y_val to (y_val - zag_depth)
			else
				set y_val to (y_val + zag_depth)
			end if
			set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
			set ctr to ctr + 1
		end repeat
		if ctr mod 2 = 0 then
			set y_val to y_val + zag_depth
			--set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
		end if
	else
		set x_val to x_val - ((x_limit - 1) * zag_length)
		set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
	end if
	
	
	set x_val to x_val - zag_depth
	if shape_width = 713 then
		set x_val to x_val - 10
	end if
	set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
	
	--left
	set ctr to 1
	if left_yn is equal to "y" or left_yn is equal to "Y" then
		repeat while ctr < y_limit
			set y_val to (y_val - zag_length)
			if ctr mod 2 = 1 then
				set x_val to (x_val + zag_depth)
			else
				set x_val to (x_val - zag_depth)
			end if
			set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
			set ctr to ctr + 1
		end repeat
	else
		set y_val to y_val - ((y_limit - 1) * zag_length)
		--end if
		set point_list to {{x_val, y_val}, {x_val, y_val}, {x_val, y_val}} & point_list
	end if
	
	
	tell canvas 1
		make new shape at beginning of graphics with properties {point list:point_list, draws shadow:false, fill:no fill, stroke color:{1.0, 0, 0.7}, thickness:1}
	end tell
	
	
end tell
