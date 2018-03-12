program importscto
	version 14.2
	
	/* Syntax: pass the filename of the XLSform
	add clear option*/
	syntax using/, Outfile(string)
	
	qui {
		/* Verify file exists */
		cap confirm file "`using'"
		if _rc {
			n di as err "XLSform not found"
			exit
		}
		
		/* Verify Excel filetype */
		cap assert regexm("`using'", ".xlsx$|.xls$")
		if _rc {
			n di as err "XLSform must have .xlsx or .xls file extension"
			exit
		}
		
		/* Verify outfile is a filepath and has .do extension */
		confirm new file `outfile'
		cap assert regexm("`outfile'", ".do$")
		if _rc {
			n di as err "The .do file extension must be used for the output file"
			exit
		}
		
		
		/* 1. Process settings tab ********************************************/
		cap import excel "`using'", sheet("settings") firstrow clear
		if _rc {
			n di as err `""settings" tab not found"'
			exit
		}
		
		cap confirm v form_title form_id version default_language
		if _rc {
			n di as err "Missing form title, ID, version, OR default language"
			exit
		}
		
		keep form_title form_id version default_language
		drop if form_id == ""

		local form_title = "`=form_title[1]'"
		local form_version = "`=version[1]'"
		local language = "`=default_language[1]'"
		local lablang = "label" + "`language'"
		
		cap file close imp
		file open imp using "`outfile'", write replace
		
		/* 2. Process choices tab *********************************************/
		cap import excel "`using'", sheet("choices") firstrow clear
		if _rc {
			n di as err `""choices" tab not found"'
			exit
		}
		
		cap confirm v list_name name
		if _rc {
			n di as err "Missing list name OR name"
			exit
		}
		
		* For surveys with multiple languages, use the labels for default language,
		* otherwise just use the single "label"
		ds label*
		if strpos("`r(varlist)'", "`lablang'") {
			local labvar `lablang'
		}
		else {
			local labvar "label"
		}

		keep list_name name `labvar'
		drop if list_name == ""
		
		* Force 32,000 character limit on value labels (probably unnecessary)
		replace `labvar' = substr(`labvar', 1, 31999)
		* Replace new line characters with space
		replace `labvar' = ustrregexra(`labvar', "\n|\r", " ")
		* Replace double quotes with single quotes
		replace `labvar' = ustrregexra(`labvar', char(34), "'")
		replace `labvar' = trim(`labvar')
		cap tostring name, replace force
		replace list_name = trim(list_name)
		replace name = trim(name)
		drop if ustrregexm(`labvar', "^\\$\{.*\}$")
		* Exit with error if labels are duplicated
		cap isid(list_name name)
		if _rc {
			n di as err "Duplicate list names in choice sheet"
			exit
		}
		
		* Destring name and drop 
		qui destring name, replace force
		drop if name == .
		sort list_name name
		
		* Generate position in each label
		bys list_name: gen resno = _n
		* Generate length of each label
		bys list_name: gen len = _N
		
		file write imp "* Defining value labels for all choices" _n(2) ///
		"#delimit ;" _n

		* Loop through each label value (each "observation" in choices sheet)
		count
		forval i = 1/`r(N)' {
			* For first label value, write "label define" command to do-file
			if `=resno[`i']' == 1 {
				file write imp `"label define `=list_name[`i']' `=name[`i']' "`=`labvar'[`i']'""' _n
			}
			
			else if `=resno[`i']' > 1 & `=resno[`i']' < `=len[`i']' {
				file write imp _char(9) `"`=name[`i']' "`=`labvar'[`i']'""' _n
			}
			
			* Add semicolon delimiter for last value in label
			else if `=resno[`i']' == `=len[`i']' {
				file write imp _char(9) `"`=name[`i']' "`=`labvar'[`i']'";"' _n(2)
			}
		}

		file write imp "#delimit cr" _n(2)
		
		/* 3. Process survey tab **********************************************/
		
		cap import excel "`using'", sheet("survey") firstrow clear
		if _rc {
			n di as err `""survey" tab not found"'
			exit
		}
		
		cap confirm v type name `labvar'
		if _rc {
			n di as err "Missing type, name, OR label"
			exit
		}
		
		drop if missing(type)
		keep type name `labvar'
		
		* Enforce Stata's 80 character limit for variable labels
		replace `labvar' = substr(`labvar', 1, 79)
		* Remove newlines
		replace `labvar' = ustrregexra(`labvar', "\n|\r", " ")
		* Replace double quotes with single quotes
		replace `labvar' = ustrregexra(`labvar', char(34), "'")
		replace `labvar' = trim(`labvar')
		cap tostring name, replace force
		replace name = trim(name)
		replace type = trim(type)
		
		* Generate indicator for repeated fields:
		gen repeated = 0
		local x = 0
		count
		forval i = 1/`r(N)' {
			if "`=type[`i']'"=="begin repeat" {
				local x = `x' + 1
			}
			else if "`=type[`i']'"=="end repeat" {
				local x = `x' - 1
			}
			
			if `x' >= 1 qui replace repeated = 1 if _n == `i'
		}
		
		* Process text and calculate fields for tostring:
		gen txt = type == "text" | type == "calculate"
		gsort -txt
		local txt
		count if txt
		forval i = 1/`r(N)' {
			if `=repeated[`i']' {
				local txt = "`txt' " + "`=name[`i']'" + "*"
			}
			
			else {
				local txt = "`txt' " + "`=name[`i']'"
			}
		}
		
		local x = 1
		foreach word in `txt' {
			if `x' == 1 { 
				file write imp "* To-string all text and calculate fields:" _n ///
				"#delimit ;" _n "local tslist `word'" _n
			}
			
			else if "`ferest()'" != "" {
				file write imp _char(9) "`word'" _n
			}	
			
			else if "`ferest()'" == "" {
				file write imp _char(9) "`word';" _n ///
				"#delimit cr" _n(2)
			}
			
			local ++x
		}

		file write imp "foreach stub in \`tslist' {" _n ///
		_char(9) "cap unab x: \`stub'" _n ///
		_char(9) "if !_rc {" _n ///
		_char(9) _char(9) "foreach var in \`x' {" _n ///
		_char(9) _char(9) _char(9) "tostring \`var', replace force" _n ///
		_char(9) _char(9) "}" _n ///
		_char(9) "}" _n ///
		_char(9) "else {" _n  ///
		_char(9) _char(9) "di " _char(34) "\`stub' do(es) not exist (yet)" _char(34) _n ///
		_char(9) "}" _n ///
		"}" _n(2)
		
		* Process integer and decimal fields to destring:
		gen num = type == "integer" | type == "decimal"
		gsort -num
		local num
		count if num
		forval i = 1/`r(N)' {
			if `=repeated[`i']' {
				local num = "`num' " + "`=name[`i']'" + "*"
			}
			
			else {
				local num = "`num' " + "`=name[`i']'"
			}
		}
		
		local x = 1
		foreach word in `num' {
			if `x' == 1 { 
				file write imp "* Destring all integer and decimal fields:" _n ///
				"#delimit ;" _n "local dslist `word'" _n
			}
			
			else if "`ferest()'" != "" {
				file write imp _char(9) "`word'" _n
			}	
			
			else if "`ferest()'" == "" {
				file write imp _char(9) "`word';" _n ///
				"#delimit cr" _n(2)
			}
			
			local ++x
		}

		file write imp "foreach stub in \`dslist' {" _n ///
		_char(9) "cap unab x: \`stub'" _n ///
		_char(9) "if !_rc {" _n ///
		_char(9) _char(9) "foreach var in \`x' {" _n ///
		_char(9) _char(9) _char(9) "destring \`var', replace force" _n ///
		_char(9) _char(9) "}" _n ///
		_char(9) "}" _n ///
		_char(9) "else {" _n  ///
		_char(9) _char(9) "di " _char(34) "\`stub' do(es) not exist (yet)" _char(34) _n ///
		_char(9) "}" _n ///
		"}" _n(2)
		
		* Process note fields to drop:
		gen nts = type == "note"
		gsort -nts
		local nts
		count if nts
		forval i = 1/`r(N)' {
			if `=repeated[`i']' {
				local nts = "`nts' " + "`=name[`i']'" + "*"
			}
			
			else {
				local nts = "`nts' " + "`=name[`i']'"
			}
		}
		
		local x = 1
		foreach word in `nts' {
			if `x' == 1 {
				file write imp "* Dropping all note fields:" _n ///
				"#delimit ;" _n "local ntlist `word'" _n
			}
			
			else if "`ferest()'" != "" {
				file write imp _char(9) "`word'" _n
			}
			
			else if "`ferest()'" == "" {
				file write imp _char(9) "`word';" _n "#delimit cr" _n(2)
			}
			local ++x
		}

		file write imp "foreach stub in \`ntlist' {" _n ///
		_char(9) "cap unab x: \`stub'" _n ///
		_char(9) "if !_rc {" _n ///
		_char(9) _char(9) "foreach var in \`x' {" _n ///
		_char(9) _char(9) _char(9) "drop \`var'" _n ///
		_char(9) _char(9) "}" _n ///
		_char(9) "}" _n ///
		_char(9) "else {" _n  ///
		_char(9) _char(9) "di " _char(34) "\`stub' do(es) not exist (yet)" _char(34) _n ///
		_char(9) "}" _n ///
		"}" _n(2)
		
		* Destring all select_one questions to value label
		gen sel = regexm(type, "select_one")
		gsort -sel
		local sel
		count if sel
		forval i = 1/`r(N)' {
			if `=repeated[`i']' {
				local sel = "`sel' " + "`=name[`i']'" + "*"
			}
			
			else {
				local sel = "`sel' " + "`=name[`i']'"
			}
		}
		
		local x = 1
		foreach word in `sel' {
			if `x' == 1 { 
				file write imp "* Destring all select_one fields for value labels:" _n ///
				"#delimit ;" _n "local solist `word'" _n
			}
			
			else if "`ferest()'" != "" {
				file write imp _char(9) "`word'" _n
			}	
			
			else if "`ferest()'" == "" {
				file write imp _char(9) "`word';" _n ///
				"#delimit cr" _n(2)
			}
			
			local ++x
		}

		file write imp "foreach stub in \`solist' {" _n ///
		_char(9) "cap unab x: \`stub'" _n ///
		_char(9) "if !_rc {" _n ///
		_char(9) _char(9) "foreach var in \`x' {" _n ///
		_char(9) _char(9) _char(9) "destring \`var', replace force" _n ///
		_char(9) _char(9) "}" _n ///
		_char(9) "}" _n ///
		_char(9) "else {" _n  ///
		_char(9) _char(9) "di " _char(34) "\`stub' do(es) not exist (yet)" _char(34) _n ///
		_char(9) "}" _n ///
		"}" _n(2)
		
		* Split select multiple questions and destring
		gen mul = regexm(type, "select_multiple")
		gsort -mul
		local mul
		count if mul
		forval i = 1/`r(N)' {
			if `=repeated[`i']' {
				local mul = "`mul' " + "`=name[`i']'" + "*"
			}
			
			else {
				local mul = "`mul' " + "`=name[`i']'"
			}
		}
		
		local x = 1
		foreach word in `mul' {
			if `x' == 1 {
				file write imp "* Split and destring all select_multiple vars:" _n ///
				"#delimit ;" _n "local smlist `word'" _n
			}
			
			else if "`ferest()'" != "" {
				file write imp _char(9) "`word'" _n
			}
			
			else if "`ferest()'" == "" {
				file write imp _char(9) "`word';" _n ///
				"#delimit cr" _n(2)
			}
			local ++x
		}

		file write imp "foreach stub in \`smlist' {" _n ///
		_char(9) "cap unab x: \`stub'" _n ///
		_char(9) "if !_rc {" _n ///
		_char(9) _char(9) "foreach var in \`x' {" _n ///
		_char(9) _char(9) _char(9) "cap confirm str v \`var'" _n ///
		_char(9) _char(9) _char(9) "if !_rc {" _n ///
		_char(9) _char(9) _char(9) _char(9) "char \`var'[multi] yes" _n ///
		_char(9) _char(9) _char(9) _char(9) "split \`var', destring force" _n ///
		_char(9) _char(9) _char(9) "}" _n ///
		_char(9) _char(9) _char(9) "else {" _n ///
		_char(9) _char(9) _char(9) _char(9) "di " _char(34) "\`var' is empty" _char(34) _n ///
		_char(9) _char(9) _char(9) "}" _n ///
		_char(9) _char(9) "}" _n ///
		_char(9) "}" _n ///
		_char(9) "else {" _n  ///
		_char(9) _char(9) "di " _char(34) "\`stub' do(es) not exist (yet)" _char(34) _n ///
		_char(9) "}" _n ///
		"}" _n(2)
		
		* Format datetime fields:
		gen dat = type == "start" | type == "end" | type == "datetime"
		gsort -dat
		local dat "SubmissionDate"
		count if dat
		forval i = 1/`r(N)' {
			if `=repeated[`i']' {
				local dat = "`dat' " + "`=name[`i']'" + "*"
			}
			
			else {
				local dat = "`dat' " + "`=name[`i']'"
			}
		}
		
		local x = 1
		foreach word in `dat' {
			if `x' == 1 {
				file write imp "* Format all date and datetime variables:" _n ///
				"#delimit ;" _n "local dtlist `word'" _n
			}
			
			else if "`ferest()'" != "" {
				file write imp _char(9) "`word'" _n
			}
			
			else if "`ferest()'" == "" {
				file write imp _char(9) "`word';" _n ///
				"#delimit cr" _n(2)
			}
			local ++x
		}	

		file write imp "foreach stub in \`dtlist' {" _n ///
		_char(9) "cap unab x: \`stub'" _n ///
		_char(9) "if !_rc {" _n ///
		_char(9) _char(9) "foreach var in \`x' {" _n ///
		_char(9) _char(9) _char(9) "gen \`var'_X = clock(\`var', " _char(34) "MDYhms" _char(34) ")" _n ///
		_char(9) _char(9) _char(9) "format \`var'_X %tc" _n ///
		_char(9) _char(9) _char(9) "drop \`var'" _n ///
		_char(9) _char(9) _char(9) "rename \`var'_X \`var'" _n ///
		_char(9) _char(9) "}" _n ///
		_char(9) "}" _n ///
		_char(9) "else {" _n  ///
		_char(9) _char(9) "di " _char(34) "\`stub' do(es) not exist (yet)" _char(34) _n ///
		_char(9) "}" _n ///
		"}" _n(2)

		drop if regexm(type, "group|repeat|note")
		gen list_name = cond(sel | mul, trim(regexr(type, "select_(one|multiple)", "")), "")
		
		gen vallab = cond(sel & !repeated, "cap label val " + name + " " + list_name, ///
		cond(sel & repeated, "cap label val " + name + "* " + list_name, ""))
		
		file write imp "* Adding value labels for all vars:" _n
		count
		forval i = 1/`r(N)' {
			if "`=vallab[`i']'" != "" {
				file write imp "`=vallab[`i']'" _n
				file write imp "if _rc == 0 {" _n
				file write imp _char(9) "order TMP, before(`=name[`i']')" _n
				file write imp _char(9) "drop `=name[`i']'" _n
				file write imp _char(9) "rename TMP `=name[`i']'" _n
				file write imp "}" _n(2)
			}
		}

		file write imp _n(2) "* Adding variable labels for non-repeated variables" _n
		count
		forval i = 1/`r(N)' {
			if `=repeated[`i']' == 0 & "`=`labvar'[`i']'" != "" {
				file write imp "cap label var `=name[`i']' " _char(34) "`=`labvar'[`i']'" _char(34) _n
			}
		}
		
		file write imp _n(2) "sort _uuid _submission_time" _n ///
		 "duplicates drop _uuid, force" _n ///
		 "qui compress _all" _n ///
		 `"save "\${dtafile}", replace"' _n(10)

		 
		file close imp
			
	}
	
end