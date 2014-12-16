<% 
    Option Explicit
	
	Dim table,part,tail_part
    ' 다중 게시판 변수
    table=request("table")
	
	
	select case table
	
	 case "asp_board"
	  part=request("table")
      tail_part="asp_tail"	 
      response.redirect "list.asp?part="&part&"&tail_part="&tail_part
	
	 case "php_board"
	  part=table
      tail_part="php_tail"
      response.redirect "list.asp?part="&part&"&tail_part="&tail_part
	
	 case "java_board"
      part=table
      tail_part="java_tail"
      response.redirect "list.asp?part="&part&"&tail_part="&tail_part
	
	 case "story_board"
	  part=table
      tail_part="story_tail"
      response.redirect "list.asp?part="&part&"&tail_part="&tail_part
	
	 case "tip_board"	
      part=table
	  tail_part="tip_tail"
      response.redirect "list.asp?part="&part&"&tail_part="&tail_part
	 
	 case "public_board"
	  part=table
      tail_part="public_tail"
      response.redirect "list.asp?part="&part&"&tail_part="&tail_part
	
	 case "job_board"
	  part=table
      tail_part="job_tail"
      response.redirect "list.asp?part="&part&"&tail_part="&tail_part
	
	 case "flash_board"
	  part=table
      tail_part="flash_tail"
      response.redirect "list.asp?part="&part&"&tail_part="&tail_part
	
	 case "html_board"
	  part=table
      tail_part="html_tail"
      response.redirect "list.asp?part="&part&"&tail_part="&tail_part
	
	end select
%>
	
