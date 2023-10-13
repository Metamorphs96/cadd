(defun ptStr2 (ptA / x1 y1)
  (setq
     x1  (car ptA)
     y1  (cadr ptA)
   )
  (strcat (rtos x1) "," (rtos y1))
)

(defun pointStr2(x y) 
	(strcat (rtos x) "," (rtos y))
)

(defun startLine(fp)
	(write-line "LINE" fp)
)

(defun StartPline(fp)
	(write-line "PLINE" fp)
)

(defun closeLines(fp)
    (write-line "C" fp)
)

(defun drawLine(fp pt1 pt2):
    (write-line (strcat "LINE " (ptStr2 ptA) " " (ptStr2 ptB) " ") fp)
)

;Send Draw Commands directly to Command Processor
(defun drawBox1 (TCnr BCnr / pt2 pt4)
   (setq pt2 (list (car BCnr) (cadr TCnr)))
   (setq pt4 (list (car TCnr) (cadr BCnr)))
   (command "pline" TCnr pt2 BCnr pt4 "c")
)

(defun drawBox4B (ptA Lx Ly / pt2 pt3 pt4)
   (setq pt2 (list (car ptA) (+ (cadr ptA) Ly) ))
   (setq pt3 (list (+ (car pt2) Lx) (cadr pt2) ))
   (setq pt4 (list (car pt3) (- (cadr pt3) Ly) ))
   (command "pline" ptA pt2 pt3 pt4 "c")
)

(defun drawBox4C (ptA Lx Ly)
   (command 
		"pline" 
	   (ptStr2 ptA) 
	   (strcat "@" (pointStr2 0 Ly)) 
	   (strcat "@" (pointStr2 Lx 0)) 
	   (strcat "@" (pointStr2 0 (- Ly))) 
		"c"
	)
)


;Save Draw Commands to File
(defun drawBox4 (fp ptA Lx Ly )
		(StartPline fp)
		(write-line (ptStr2 ptA) fp)
		(write-line (strcat "@" (pointStr2 0 Ly)) fp)
		(write-line (strcat "@" (pointStr2 Lx 0)) fp)
		(write-line (strcat "@" (pointStr2 0 (- Ly))) fp)
		(closeLines fp)
)

(defun ScriptWriter1 (scrFile / ptA Lx Ly)
	(setq 
		ptA (list 0 0)
		Ly 16
		Lx (* 2 Ly)
	)
	(progn	
		(drawBox4 scrFile ptA Lx Ly)
		(write-line "ZOOM E" scrFile)
		;(drawBox4C ptA Lx Ly)
	)
)

(defun ScriptWriter2 (scrFile / ptA Lx Ly)
	(setq 
	    pt0 (list 0 0)
		ptA (list 0 0)
		BuildingHeight 2.4
		BuildingWidth 8
		BuildingLength (* 2 BuildingWidth)
	)
	(progn
		;Plan
		(drawBox4 scrFile ptA BuildingLength BuildingWidth)
		
		;Elevation 1
		(setq ptA (list (car ptA) (- (cadr ptA) BuildingHeight) ))
		(drawBox4 scrFile ptA BuildingLength BuildingHeight)
		
		;Elevation 2
		(setq ptA (list (+ (car ptA) BuildingLength) (cadr ptA)  ))
		(drawBox4 scrFile ptA BuildingWidth BuildingHeight)
		
		;Elevation 3
		(setq ptA (list (- (car pt0) BuildingWidth) (cadr ptA) ))
		(drawBox4 scrFile ptA BuildingWidth BuildingHeight)
		
		;Elevation 4
		(setq ptA (list (- (car ptA) BuildingLength) (cadr ptA) ))
		(drawBox4 scrFile ptA BuildingLength BuildingHeight)
		
		;Section
		(setq ptA (list (+ (car pt0) BuildingLength BuildingWidth) (cadr ptA) ))
		(drawBox4 scrFile ptA BuildingWidth BuildingHeight)	
		
		(write-line "ZOOM E" scrFile)

	)
)







(defun C:CmdMain ()
  (setq
     fn (strcat (getvar "DWGPREFIX") "DrawBox2a.scr")
  );end setq
  
  ;write output file
  (If (setq scrFile (open fn "w") )
    (progn
     (ScriptWriter2 scrFile)
     (close scrFile)
    );progn
  ;else
    (write-line "output file NOT found")
  );endif
)