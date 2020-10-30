open LessPainfulGoogleSheets
open LessPainfulXlsx

let getAssignments () =
    let getDisciplineCode (assignment: string) =
        assignment.Split ' ' |> Seq.head

    let cleanTeacher (teacher: string) = 
        teacher.Split ',' |> Seq.head

    let assignments = openXlsxSheet "assignments.xlsx" 0

    let assignmentColumn = readColumnByName assignments "Педагогическое задание" |> Seq.toList
    let assignmentKindColumn = readColumnByName assignments "Виды учебной работы" |> Seq.toList
    let teacherColumn = readColumnByName assignments "Преподаватель" |> Seq.toList
    let departmentColumn = readColumnByName assignments "SAP-подразделение 2" |> Seq.toList

    let mutable result = Map.empty

    let rec traverse assignmentColumn assignmentKindColumn teacherColumn departmentColumn = 
        if assignmentColumn |> Seq.isEmpty |> not then 
            let assignmentCode = assignmentColumn |> Seq.head |> getDisciplineCode
            let kind = assignmentKindColumn |> Seq.head
            let teacher = teacherColumn |> Seq.head |> cleanTeacher
            let department = departmentColumn |> Seq.head

            if result.ContainsKey assignmentCode then
                if kind = "Лекции" then
                    result <- result.Add (assignmentCode, (teacher, department))
            else
                result <- result.Add (assignmentCode, (teacher, department))

            traverse (assignmentColumn |> Seq.tail) (assignmentKindColumn |> Seq.tail) (teacherColumn |> Seq.tail) (departmentColumn |> Seq.tail)
    
    traverse assignmentColumn assignmentKindColumn teacherColumn departmentColumn

    result

[<EntryPoint>]
let main argv =
    let getDisciplineCode (discipline: string) =
        discipline.Split ' ' |> Seq.head |> (fun (s: string) -> s.Trim [|'['; ']'|])

    printfn "Getting assignments from .xlsx..."

    let assignments = getAssignments ()

    printfn "Connecting to Google Sheets..."

    let spreadsheetId = "1TsOCZltVSHxQouOTU8WDyW4oxITs1xUyV_GC-TfqbfI"
    let page = "2017"
    use service = openGoogleSheet "AssignmentMatcher" 

    printfn "Reading disciplines..."

    let disciplineNames = readGoogleSheetColumn service spreadsheetId page "A" 2

    let mutable teachers = []
    let mutable departments = []

    for discipline in disciplineNames do
        let code = getDisciplineCode discipline
        if assignments.ContainsKey code then 
            let teacher, department = assignments.[code]
            teachers <- teacher :: teachers
            departments <- department :: departments
        else
            printfn "Warning! Discipline %s has no assingment!" discipline
            teachers <- "" :: teachers
            departments <- "" :: departments

    teachers <- List.rev teachers    
    departments <- List.rev departments

    printfn "Writing changes..."

    writeGoogleSheetColumn service spreadsheetId page "B" 2 teachers
    writeGoogleSheetColumn service spreadsheetId page "C" 2 departments

    0
