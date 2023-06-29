option explicit
dim objShell, objFile
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
dim directory
directory = objFile.GetParentFolderName(wscript.ScriptFullName)

sub Main()
    'Debug()

    dim i, j, element, data, path, test, chartType, difficulty, level, pointer, levelData, bpms, maidata, outputPath
    outputPath = directory & "\output.txt"
    path = ""
    'path = directory & "\test\Silly Love.sm"
    'path = directory & "\test2\Idola.sm"
    'path = directory & "\test3\666.sm"
    'path = directory & "\DEAD BEATS.sm"
    path = directory & "\I.sm"

    select case wscript.Arguments.length
        case 0
            if path = "" then
                path = inputbox("Enter path of simfile (.sm)")
            end if

            if path = "" then
                wscript.Quit()
            end if
        case 1
            path = wscript.Arguments(0)
        case 2
            path = wscript.Arguments(0)
            outputPath = wscript.Arguments(1)
        case else
            wscript.Echo("Invalid arguments")
            wscript.Quit()
    end select

    data = objFile.OpenTextFile(path).ReadAll()
    data = Strip(data)
    pointer = 0
    data = split(data, vbcrlf)

    for i = 0 to ubound(data)
        data(i) = Strip(data(i))
    next

    for i = 0 to ubound(data)
        if left(data(i), 6) = "#BPMS:" then
            pointer = i
            exit for
        end if

        if i = ubound(data) then
            wscript.Echo("Could not find bpm")
            wscript.Quit()
        end if
    next

    for i = pointer to ubound(data)
        if right(data(i), 1) = ";" then
            bpms = bpms & left(data(i), len(data(i)) - 1) & vbcrlf
            exit for
        end if

        bpms = bpms & data(i) & vbcrlf

        if i = ubound(data) then
            wscript.Echo("Something went wrong")
            wscript.Quit()
        end if
    next

    bpms = mid(bpms, 7)
    bpms = split(bpms, ",")

    for i = 0 to ubound(bpms)
        bpms(i) = Strip(bpms(i))
    next

    chartType = inputbox("Enter chart type (0: pump-single, 1: pump-double)")

    if chartType = "" then
        wscript.Quit()
    end if

    select case chartType
        case 0
            chartType = "pump-single:"
        case 1
            chartType = "pump-double:"
        case else
            wscript.Echo("Invalid input")
            Main()
            exit sub
    end select

    for i = 0 to ubound(data)
        if data(i) = chartType then
            pointer = i
            exit for
        end if

        if i = ubound(data) then
            wscript.Echo("Chart type not found")
            Main()
            exit sub
        end if
    next

    difficulty = inputbox("Enter chart difficulty (0: Beginner, 1: Easy, 2: Medium, 3: Hard, 4: Challenge, 5: Edit)")
    
    if difficulty = "" then
        wscript.Quit()
    end if

    select case difficulty
        case 0
            difficulty = "Beginner:"
        case 1
            difficulty = "Easy:"
        case 2
            difficulty = "Medium:"
        case 3
            difficulty = "Hard:"
        case 4
            difficulty = "Challenge:"
        case 5
            difficulty = "Edit:"
        case else
            wscript.Echo("Invalid input")
            Main()
            exit sub
    end select

    for i = pointer to ubound(data)
        if data(i) = difficulty then
            pointer = i
            exit for
        end if

        if i = ubound(data) then
            wscript.Echo("Difficulty not found")
            Main()
            exit sub
        end if
    next
    
    level = inputbox("Enter chart level (n > 0)")

    if level = "" then
        wscript.Quit()
    end if

    if cint(level) <= 0 then
        wscript.Echo("Invalid input")
        Main()
        exit sub
    end if

    level = level & ":"

    for i = pointer to pointer + 5
        if data(i) = level then
            pointer = i
            exit for
        end if

        if i = pointer + 5 then
            wscript.Echo("Level not found")
            Main()
            exit sub
        end if
    next

    for i = pointer to ubound(data)
        select case chartType
            case "pump-single:"
                if data(i) = "00000" then
                    pointer = i
                    exit for
                end if
            case "pump-double:"
                if data(i) = "0000000000" then
                    pointer = i
                    exit for
                end if
        end select

        if i = ubound(data) then
            wscript.Echo("An error occured")
            Main()
            exit sub
        end if
    next

    levelData = ""

    for i = pointer to ubound(data)
        element = data(i)

        if instr(element, "//") then
            element = left(element, instr(element, "//") - 1)
        end if

        levelData = levelData & element & vbcrlf

        if left(data(i), 1) = ";" then
            exit for
        end if
    next

    levelData = Strip(levelData)
    levelData = left(levelData, len(levelData) - 1)
    levelData = Strip(levelData)
    levelData = split(levelData, ",")

    for i = 0 to ubound(levelData)
        levelData(i) = Strip(levelData(i))
    next

    select case split(levelData(0), vbcrlf)(0)
        case "00000"
            maidata = ConvertSimfileToMaidata(levelData, bpms)
        case "0000000000"
            maidata = ConvertSimfileToMaidataDouble(levelData, bpms)
        case else
            wscript.Echo("Something went wrong")
            wscript.Quit()
    end select
    
    wscript.Echo("Successfully exported to " & outputPath)
    objFile.CreateTextFile(outputPath, true).Write(maidata)
end sub

function ConvertSimfileToMaidata(levelData, bpms)
    dim i, j, k, element, measures, maxMeasure, length, position, measureData, tempMaidata, maidata, currentBpm, currentMeasure, pointer, count, lastPosition, bpmsPointer, currentBeat, bpmItem

    measures = array()

    for each element in levelData
        measures = PushArray(measures, ubound(split(element, vbcrlf)) + 1)
    next

    maxMeasure = 0

    for each element in measures
        if element > maxMeasure then
            maxMeasure = element
        end if
    next

    currentBpm = split(bpms(0), "=")(1)
    currentBpm = eval(currentBpm)
    bpmsPointer = 1

    if bpmsPointer <= ubound(bpms) then
        bpmItem = split(bpms(bpmsPointer), "=")
    end if

    maidata = "(" & currentBpm & ")" & vbcrlf
    currentMeasure = 0
    lastPosition = "ul"
    position = -1
    currentBeat = 0

    for i = 0 to ubound(levelData)
        measureData = split(levelData(i), vbcrlf)
        wscript.Echo("Progress: " & i & " / " & ubound(levelData))

        if currentMeasure = measures(i) then
        else
            currentMeasure = measures(i)
            maidata = maidata & "{" & currentMeasure & "}" & vbcrlf
        end if

        for j = 0 to ubound(measureData)
            measureData(j) = Strip(measureData(j))
            position = position + 1
            count = 0

            for k = 1 to len(measureData(j))
                if mid(measureData(j), k, 1) = "1" or mid(measureData(j), k, 1) = "2" then
                    count = count + 1
                end if
            next

            select case count
                case 0
                    maidata = maidata & "," & vbcrlf
                case 1
                    if mid(measureData(j), 1, 1) = "1" then
                        maidata = maidata & "6," & vbcrlf
                        lastPosition = "dl"
                    elseif mid(measureData(j), 2, 1) = "1" then
                        maidata = maidata & "7," & vbcrlf
                        lastPosition = "ul"
                    elseif mid(measureData(j), 3, 1) = "1" then
                        select case lastPosition
                            case "ul"
                                maidata = maidata & "1," & vbcrlf
                            case "ur"
                                maidata = maidata & "8," & vbcrlf
                            case "dl"
                                maidata = maidata & "4," & vbcrlf
                            case "dr"
                                maidata = maidata & "5," & vbcrlf
                        end select
                    elseif mid(measureData(j), 4, 1) = "1" then
                        maidata = maidata & "2," & vbcrlf
                        lastPosition = "ur"
                    elseif mid(measureData(j), 5, 1) = "1" then
                        maidata = maidata & "3," & vbcrlf
                        lastPosition = "dr"
                    end if

                    if mid(measureData(j), 1, 1) = "2" then
                        length = CalcSlider(levelData, position, 1)
                        maidata = maidata & "6h[" & length(1) & ":" & length(0) & "]," & vbcrlf
                        lastPosition = "dl"
                    elseif mid(measureData(j), 2, 1) = "2" then
                        length = CalcSlider(levelData, position, 2)
                        maidata = maidata & "7h[" & length(1) & ":" & length(0) & "]," & vbcrlf
                        lastPosition = "ul"
                    elseif mid(measureData(j), 3, 1) = "2" then
                        select case lastPosition
                            case "ul"
                                length = CalcSlider(levelData, position, 3)
                                maidata = maidata & "1h[" & length(1) & ":" & length(0) & "]," & vbcrlf
                            case "ur"
                                length = CalcSlider(levelData, position, 3)
                                maidata = maidata & "8h[" & length(1) & ":" & length(0) & "]," & vbcrlf
                            case "dl"
                                length = CalcSlider(levelData, position, 3)
                                maidata = maidata & "4h[" & length(1) & ":" & length(0) & "]," & vbcrlf
                            case "dr"
                                length = CalcSlider(levelData, position, 3)
                                maidata = maidata & "5h[" & length(1) & ":" & length(0) & "]," & vbcrlf
                        end select
                    elseif mid(measureData(j), 4, 1) = "2" then
                        length = CalcSlider(levelData, position, 4)
                        maidata = maidata & "2h[" & length(1) & ":" & length(0) & "]," & vbcrlf
                        lastPosition = "ur"
                    elseif mid(measureData(j), 5, 1) = "2" then
                        length = CalcSlider(levelData, position, 5)
                        maidata = maidata & "3h[" & length(1) & ":" & length(0) & "]," & vbcrlf
                        lastPosition = "dr"
                    end if
                case 3
                    tempMaidata = ""

                    if mid(measureData(j), 3, 1) = "1" or mid(measureData(j), 3, 1) = "2" then
                        if mid(measureData(j), 1, 1) = "1" then
                            tempMaidata = tempMaidata & "/5" & vbcrlf
                            lastPosition = "dl"
                        end if
                        
                        if mid(measureData(j), 2, 1) = "1" then
                            tempMaidata = tempMaidata & "/8" & vbcrlf
                            lastPosition = "ul"
                        end if
                        
                        if mid(measureData(j), 4, 1) = "1" then
                            tempMaidata = tempMaidata & "/1" & vbcrlf
                            lastPosition = "ur"
                        end if
                        
                        if mid(measureData(j), 5, 1) = "1" then
                            tempMaidata = tempMaidata & "/4" & vbcrlf
                            lastPosition = "dr"
                        end if

                        if mid(measureData(j), 1, 1) = "2" then
                            length = CalcSlider(levelData, position, 1)
                            tempMaidata = tempMaidata & "/5h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            lastPosition = "dl"
                        end if
                        
                        if mid(measureData(j), 2, 1) = "2" then
                            length = CalcSlider(levelData, position, 2)
                            tempMaidata = tempMaidata & "/8h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            lastPosition = "ul"
                        end if
                        
                        if mid(measureData(j), 4, 1) = "2" then
                            length = CalcSlider(levelData, position, 4)
                            tempMaidata = tempMaidata & "/1h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            lastPosition = "ur"
                        end if
                        
                        if mid(measureData(j), 5, 1) = "2" then
                            length = CalcSlider(levelData, position, 5)
                            tempMaidata = tempMaidata & "/4h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            lastPosition = "dr"
                        end if
                    else
                        if mid(measureData(j), 1, 1) = "1" then
                            tempMaidata = tempMaidata & "/6" & vbcrlf
                            lastPosition = "dl"
                        end if
                        
                        if mid(measureData(j), 2, 1) = "1" then
                            tempMaidata = tempMaidata & "/7" & vbcrlf
                            lastPosition = "ul"
                        end if
                        
                        if mid(measureData(j), 4, 1) = "1" then
                            tempMaidata = tempMaidata & "/2" & vbcrlf
                            lastPosition = "ur"
                        end if
                        
                        if mid(measureData(j), 5, 1) = "1" then
                            tempMaidata = tempMaidata & "/3" & vbcrlf
                            lastPosition = "dr"
                        end if

                        if mid(measureData(j), 1, 1) = "2" then
                            length = CalcSlider(levelData, position, 1)
                            tempMaidata = tempMaidata & "/6h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            lastPosition = "dl"
                        end if
                        
                        if mid(measureData(j), 2, 1) = "2" then
                            length = CalcSlider(levelData, position, 2)
                            tempMaidata = tempMaidata & "/7h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            lastPosition = "ul"
                        end if
                        
                        if mid(measureData(j), 4, 1) = "2" then
                            length = CalcSlider(levelData, position, 4)
                            tempMaidata = tempMaidata & "/2h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            lastPosition = "ur"
                        end if
                        
                        if mid(measureData(j), 5, 1) = "2" then
                            length = CalcSlider(levelData, position, 5)
                            tempMaidata = tempMaidata & "/3h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            lastPosition = "dr"
                        end if
                    end if

                    tempMaidata = mid(tempMaidata, 2)
                    maidata = maidata & tempMaidata & "," & vbcrlf
                case 5
                    maidata = maidata & "1`2`3`4`5`6`7`8," & vbcrlf
                case else
                    tempMaidata = ""

                    if mid(measureData(j), 1, 1) = "1" then
                        tempMaidata = tempMaidata & "/6" & vbcrlf
                        lastPosition = "dl"
                    end if
                    
                    if mid(measureData(j), 2, 1) = "1" then
                        tempMaidata = tempMaidata & "/7" & vbcrlf
                        lastPosition = "ul"
                    end if
                    
                    if mid(measureData(j), 4, 1) = "1" then
                        tempMaidata = tempMaidata & "/2" & vbcrlf
                        lastPosition = "ur"
                    end if
                    
                    if mid(measureData(j), 5, 1) = "1" then
                        tempMaidata = tempMaidata & "/3" & vbcrlf
                        lastPosition = "dr"
                    end if

                    if mid(measureData(j), 3, 1) = "1" then
                        select case lastPosition
                            case "ul"
                                tempMaidata = tempMaidata & "/1" & vbcrlf
                            case "ur"
                                tempMaidata = tempMaidata & "/8" & vbcrlf
                            case "dl"
                                tempMaidata = tempMaidata & "/4" & vbcrlf
                            case "dr"
                                tempMaidata = tempMaidata & "/5" & vbcrlf
                        end select
                    end if

                    if mid(measureData(j), 1, 1) = "2" then
                        length = CalcSlider(levelData, position, 1)
                        tempMaidata = tempMaidata & "/6h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                        lastPosition = "dl"
                    end if
                    
                    if mid(measureData(j), 2, 1) = "2" then
                        length = CalcSlider(levelData, position, 2)
                        tempMaidata = tempMaidata & "/7h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                        lastPosition = "ul"
                    end if
                    
                    if mid(measureData(j), 4, 1) = "2" then
                        length = CalcSlider(levelData, position, 4)
                        tempMaidata = tempMaidata & "/2h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                        lastPosition = "ur"
                    end if
                    
                    if mid(measureData(j), 5, 1) = "2" then
                        length = CalcSlider(levelData, position, 5)
                        tempMaidata = tempMaidata & "/3h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                        lastPosition = "dr"
                    end if

                    if mid(measureData(j), 3, 1) = "2" then
                        select case lastPosition
                            case "ul"
                                length = CalcSlider(levelData, position, 3)
                                tempMaidata = tempMaidata & "/1h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            case "ur"
                                length = CalcSlider(levelData, position, 3)
                                tempMaidata = tempMaidata & "/8h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            case "dl"
                                length = CalcSlider(levelData, position, 3)
                                tempMaidata = tempMaidata & "/4h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                            case "dr"
                                length = CalcSlider(levelData, position, 3)
                                tempMaidata = tempMaidata & "/5h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                        end select
                    end if

                    tempMaidata = mid(tempMaidata, 2)
                    maidata = maidata & tempMaidata & "," & vbcrlf
            end select

            currentBeat = currentBeat + (4 / eval(measures(i)))

            if bpmsPointer <= ubound(bpms) then
                if currentBeat >= eval(bpmItem(0)) - 0.001 then
                    maidata = maidata & "(" & bpmItem(1) & ")" & vbcrlf
                    bpmsPointer = bpmsPointer + 1

                    if bpmsPointer <= ubound(bpms) then
                        bpmItem = split(bpms(bpmsPointer), "=")
                    end if
                end if
            end if
        next
    next

    ConvertSimfileToMaidata = FormatMaidata(maidata)
end function

function ConvertSimfileToMaidataDouble(levelData, bpms)
    dim i, j, element, measures, currentBpm, bpmsPointer, maidata, position, lastPosition, currentBeat, currentMeasure, measureData, tempMaidata, length

    measures = array()

    for each element in levelData
        measures = PushArray(measures, ubound(split(element, vbcrlf)) + 1)
    next

    currentBpm = split(bpms(0), "=")(1)
    currentBpm = eval(currentBpm)
    bpmsPointer = 1

    if bpmsPointer <= ubound(bpms) then
        bpmItem = split(bpms(bpmsPointer), "=")
    end if

    maidata = "(" & currentBpm & ")" & vbcrlf
    currentMeasure = 0
    lastPosition = "ul"
    position = -1
    currentBeat = 0

    for i = 0 to ubound(levelData)
        measureData = split(levelData(i), vbcrlf)
        wscript.Echo("Progress: " & i & " / " & ubound(levelData))

        if currentMeasure = measures(i) then
        else
            currentMeasure = measures(i)
            maidata = maidata & "{" & currentMeasure & "}" & vbcrlf
        end if

        for j = 0 to ubound(measureData)
            measureData(j) = Strip(measureData(j))
            position = position + 1

            tempMaidata = ""

            if mid(measureData(j), 1, 1) = "1" then
                tempMaidata = tempMaidata & "/6" & vbcrlf
                lastPosition = "dl"
            end if
            
            if mid(measureData(j), 2, 1) = "1" then
                tempMaidata = tempMaidata & "/7" & vbcrlf
                lastPosition = "ul"
            end if
            
            if mid(measureData(j), 4, 1) = "1" then
                tempMaidata = tempMaidata & "/8" & vbcrlf
                lastPosition = "ur"
            end if
            
            if mid(measureData(j), 5, 1) = "1" then
                tempMaidata = tempMaidata & "/5" & vbcrlf
                lastPosition = "dr"
            end if

            if mid(measureData(j), 3, 1) = "1" then
                if mid(measureData(j), 8, 1) = "1" then
                    tempMaidata = tempMaidata & "/C" & vbcrlf
                else
                    select case lastPosition
                        case "ul"
                            tempMaidata = tempMaidata & "/B5" & vbcrlf
                        case "ur"
                            tempMaidata = tempMaidata & "/B6" & vbcrlf
                        case "dl"
                            tempMaidata = tempMaidata & "/B8" & vbcrlf
                        case "dr"
                            tempMaidata = tempMaidata & "/B7" & vbcrlf
                    end select
                end if
            end if

            if mid(measureData(j), 6, 1) = "1" then
                tempMaidata = tempMaidata & "/4" & vbcrlf
                lastPosition = "dl"
            end if
            
            if mid(measureData(j), 7, 1) = "1" then
                tempMaidata = tempMaidata & "/1" & vbcrlf
                lastPosition = "ul"
            end if
            
            if mid(measureData(j), 9, 1) = "1" then
                tempMaidata = tempMaidata & "/2" & vbcrlf
                lastPosition = "ur"
            end if
            
            if mid(measureData(j), 10, 1) = "1" then
                tempMaidata = tempMaidata & "/3" & vbcrlf
                lastPosition = "dr"
            end if

            if mid(measureData(j), 8, 1) = "1" then
                if mid(measureData(j), 3, 1) = "1" then
                else
                    select case lastPosition
                        case "ul"
                            tempMaidata = tempMaidata & "/B3" & vbcrlf
                        case "ur"
                            tempMaidata = tempMaidata & "/B4" & vbcrlf
                        case "dl"
                            tempMaidata = tempMaidata & "/B2" & vbcrlf
                        case "dr"
                            tempMaidata = tempMaidata & "/B1" & vbcrlf
                    end select
                end if
            end if

            if mid(measureData(j), 1, 1) = "2" then
                length = CalcSlider(levelData, position, 1)
                tempMaidata = tempMaidata & "/6h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                lastPosition = "dl"
            end if
            
            if mid(measureData(j), 2, 1) = "2" then
                length = CalcSlider(levelData, position, 2)
                tempMaidata = tempMaidata & "/7h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                lastPosition = "ul"
            end if
            
            if mid(measureData(j), 4, 1) = "2" then
                length = CalcSlider(levelData, position, 4)
                tempMaidata = tempMaidata & "/8h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                lastPosition = "ur"
            end if
            
            if mid(measureData(j), 5, 1) = "2" then
                length = CalcSlider(levelData, position, 5)
                tempMaidata = tempMaidata & "/5h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                lastPosition = "dr"
            end if

            if mid(measureData(j), 3, 1) = "2" then
                if mid(measureData(j), 8, 1) = "2" then
                    tempMaidata = tempMaidata & "/C" & vbcrlf
                else
                    select case lastPosition
                        case "ul"
                            tempMaidata = tempMaidata & "/B5" & vbcrlf
                        case "ur"
                            tempMaidata = tempMaidata & "/B6" & vbcrlf
                        case "dl"
                            tempMaidata = tempMaidata & "/B8" & vbcrlf
                        case "dr"
                            tempMaidata = tempMaidata & "/B7" & vbcrlf
                    end select
                end if
            end if

            if mid(measureData(j), 6, 1) = "2" then
                length = CalcSlider(levelData, position, 6)
                tempMaidata = tempMaidata & "/6h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                lastPosition = "dl"
            end if
            
            if mid(measureData(j), 7, 1) = "2" then
                length = CalcSlider(levelData, position, 7)
                tempMaidata = tempMaidata & "/7h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                lastPosition = "ul"
            end if
            
            if mid(measureData(j), 9, 1) = "2" then
                length = CalcSlider(levelData, position, 9)
                tempMaidata = tempMaidata & "/8h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                lastPosition = "ur"
            end if
            
            if mid(measureData(j), 10, 1) = "2" then
                length = CalcSlider(levelData, position, 10)
                tempMaidata = tempMaidata & "/5h[" & length(1) & ":" & length(0) & "]" & vbcrlf
                lastPosition = "dr"
            end if

            if mid(measureData(j), 8, 1) = "2" then
                if mid(measureData(j), 3, 1) = "2" then
                    tempMaidata = tempMaidata & "/C" & vbcrlf
                else
                    select case lastPosition
                        case "ul"
                            tempMaidata = tempMaidata & "/B3" & vbcrlf
                        case "ur"
                            tempMaidata = tempMaidata & "/B4" & vbcrlf
                        case "dl"
                            tempMaidata = tempMaidata & "/B2" & vbcrlf
                        case "dr"
                            tempMaidata = tempMaidata & "/B1" & vbcrlf
                    end select
                end if
            end if

            tempMaidata = mid(tempMaidata, 2)
            maidata = maidata & tempMaidata & "," & vbcrlf

            currentBeat = currentBeat + (4 / eval(measures(i)))

            if bpmsPointer <= ubound(bpms) then
                if currentBeat >= eval(bpmItem(0)) - 0.001 then
                    maidata = maidata & "(" & bpmItem(1) & ")" & vbcrlf
                    bpmsPointer = bpmsPointer + 1

                    if bpmsPointer <= ubound(bpms) then
                        bpmItem = split(bpms(bpmsPointer), "=")
                    end if
                end if
            end if
        next
    next

    ConvertSimfileToMaidataDouble = FormatMaidata(maidata)
end function

function CalcSlider(byval levelData, position, step)
    dim i, element, measures, timeline, length, temp
    measures = array()

    for each element in levelData
        measures = PushArray(measures, ubound(split(element, vbcrlf)) + 1)
    next

    levelData = join(levelData, vbcrlf)
    levelData = split(levelData, vbcrlf)
    timeline = array()

    for each element in measures
        for i = 1 to element
            timeline = PushArray(timeline, element)
        next
    next

    length = array(0, 1)

    for i = position to ubound(timeline)
        if mid(levelData(i), step, 1) = "3" then
            exit for
        end if

        temp = CalcFraction(length(0), length(1), "+", 1, cint(timeline(i)))
        length = temp

        if i = ubound(timeline) then
            Breakpoint("Something went wrong")
        end if
    next

    CalcSlider = length
end function

function FormatMaidata(maidata)
    dim i, element, result, count
    maidata = replace(maidata, vbcrlf, "")
    result = ""
    count = 0

    for i = 1 to len(maidata)
        element = mid(maidata, i, 1)

        select case element
            case "("
                result = result & vbcrlf & vbcrlf & element
                count = 0
            case ")"
                result = result & element & vbcrlf
                count = 0
            case "{"
                result = result & vbcrlf & vbcrlf & element
                count = 0
            case "}"
                result = result & element & vbcrlf
                count = 0
            case ","
                count = count + 1
                result = result & element
            case else
                if count > 1 then
                    result = result & vbcrlf & element
                else
                    result = result & element
                end if

                count = 0
        end select
    next

    FormatMaidata = result
end function

function FindIndex(input, find)
    dim i

    for i = 0 to ubound(input)
        if input(i) = find then
            FindIndex = i
            exit function
        end if
    next

    FindIndex = -1
end function

function PushArray(input, push)
    dim temp, i
    redim temp(ubound(input) + 1)

    for i = 0 to ubound(input)
        temp(i) = input(i)
    next

    temp(ubound(temp)) = push
    PushArray = temp
end function

function Strip(byval input)
    dim i, stage, clean
    clean = true
    stage = 0

    for i = 0 to 999
        input = trim(input)

        select case stage
            case 0
                if left(input, 1) = vbcr then
                    clean = false
                    input = mid(input, 2)
                else
                    stage = 1
                end if
            case 1
                if right(input, 1) = vbcr then
                    clean = false
                    input = left(input, len(input) - 1)
                else
                    stage = 2
                end if
            case 2
                if left(input, 1) = vblf then
                    clean = false
                    input = mid(input, 2)
                else
                    stage = 3
                end if
            case 3
                if right(input, 1) = vblf then
                    clean = false
                    input = left(input, len(input) - 1)
                else
                    if clean = true then
                        Strip = input
                        exit function
                    else
                        clean = true
                        stage = 0
                    end if
                end if
        end select

        if i = 999 then
            wscript.Echo("Overload limit reached")
            Strip = false
            exit function
        end if
    next
end function

function CalcFraction(byval num1, byval den1, opr, byval num2, byval den2)
    dim num3, den3, gcf
    num1 = int(num1)
    num2 = int(num2)
    den1 = int(den1)
    den2 = int(den2)
    den3 = LeastCommonMultiple(den1, den2)
    num1 = num1 * (den3 / den1)
    num2 = num2 * (den3 / den2)

    select case opr
        case "+"
            num3 = num1 + num2
        case "-"
            num3 = num1 - num2
        case "*"
            num3 = num1 * num2
            den3 = den3 * den3
        case "/"
            num3 = num1 * den3
            den3 = den3 * num2
        case else
            wscript.Echo("Invalid operator")
            CalcFraction = false
            exit function
    end select

    gcf = GreatestCommonFactor(num3, den3)
    num3 = num3 / gcf
    den3 = den3 / gcf

    if den3 < 0 then
        num3 = num3 * -1
        den3 = den3 * -1
    end if

    CalcFraction = array(num3, den3)
end function

function LeastCommonMultiple(byval num1, byval num2)
    dim i, element, element2, mul1, mul2
    mul1 = array()
    mul2 = array()
    num1 = abs(num1)
    num2 = abs(num2)

    for i = 1 to 999
        mul1 = PushArray(mul1, num1 * i)
        mul2 = PushArray(mul2, num2 * i)

        for each element in mul1
            for each element2 in mul2
                if element = element2 then
                    LeastCommonMultiple = element
                    exit function
                end if
            next
        next
    next

    wscript.Echo("Loop limit exceeded")
    LeastCommonMultiple = false
end function

function GreatestCommonFactor(byval num1, byval num2)
    dim i, fac

    num1 = abs(num1)
    num2 = abs(num2)

    if num1 > num2 then
        fac = num2
    else
        fac = num1
    end if

    for i = 0 to 999
        if (num1 / fac) - int(num1 / fac) = 0 then
            if (num2 / fac) - int(num2 / fac) = 0 then
                GreatestCommonFactor = fac
                exit function
            end if
        end if

        fac = fac - 1

        if fac = 0 then
            wscript.Echo("Something went wrong")
            GreatestCommonFactor = false
            exit function
        end if
    next

    wscript.Echo("Loop limit exceeded")
end function

sub Breakpoint(message)
    if typeName(message) = "Variant()" then
        wscript.Echo("[" & join(message, ", ") & "]")
        wscript.Quit()
    end if

    wscript.Echo(message)
    wscript.Quit()
end sub

sub Debug()
    dim test
    test = CalcFraction(3,4,"/",4,5)
    Breakpoint(test)

    wscript.Quit()
end sub

Main()
