DIM arr$(110000)
OPEN "G:\ICS4\schedule\scheduler\Newsched.txt" FOR INPUT AS #1
OPEN "G:\ICS4\schedule\scheduler\bigboi.txt" FOR OUTPUT AS #2
count = 0
FOR i = 0 TO 103470
    PRINT i
    IF count = 0 THEN
        INPUT #1, person$
    ELSE
        INPUT #1, temp$
    END IF
    IF temp$ = person$ THEN
        count = count + 1
    ELSE

        PRINT #2, person$

        person$ = temp$
        FOR j = 1 TO count
            PRINT #2, arr$(j)
        NEXT
        count = 1
    END IF
    IF count < 8 THEN
        INPUT #1, arr$(i)
    ELSE
        INPUT #1, garb
    END IF
NEXT
CLOSE #2
CLOSE #1
