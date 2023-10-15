;; Common Lisp version of program


(defun ptStr2 (ptA)
  (setq
     x1  (car ptA)
     y1  (cadr ptA)
   )
  (concatenate 'String (write-to-string x1) "," (write-to-string y1))
)

(defun pointStr2(x y) 
	(concatenate 'String (write-to-string x) "," (write-to-string y))
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
    (write-line (concatenate 'String  "LINE " (ptStr2 ptA) " " (ptStr2 ptB) " ") fp)
)


;Save Draw Commands to File
(defun drawBox4 (fp ptA Lx Ly )
		(StartPline fp)
		(write-line (ptStr2 ptA) fp)
		(write-line (concatenate 'String  "@" (pointStr2 0 Ly)) fp)
		(write-line (concatenate 'String  "@" (pointStr2 Lx 0)) fp)
		(write-line (concatenate 'String  "@" (pointStr2 0 (- Ly))) fp)
		(closeLines fp)
)

(defun ScriptWriter1 (scrFile)
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

(defun ScriptWriter2 (scrFile)
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






;;Don't have access to drawing path DWGPREFIX
;;So cannot use to contruct location for file path
;;Therefore for simplicity just hardcode path
;;Determine how to construct paths at later date
(defun CmdMain ()
  (setq
     fn (concatenate 'String "c:/yourpath/" "clDrawBox2.scr")
  );end setq
  
  ;write output file
  (If (setq scrFile (open fn :direction :output :if-exists :supersede) )
    (progn
     (ScriptWriter2 scrFile)
     (close scrFile)
    );progn
  ;else
    (write-line "output file NOT found")
  );endif
)