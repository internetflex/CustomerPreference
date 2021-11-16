open System
open System.IO
open FSharp.Interop.Excel

type CustomerDetailsType = ExcelFile<"CustomerData.xlsx", SheetName="Customer Info", HasHeaders=true>

type Preference =
    | Day of int
    | WeekDay of int list
    | EveryDay
    | Never

type CustomerDetails =
    {
        CustomerName : string option
        Preference : Preference option
    }

    static member Default =
    {
        CustomerName = None
        Preference = None
    }

let getCustomerName (row:CustomerDetailsType.Row) (customerDetails: CustomerDetails) =
    match row.``Customer Name`` with
    | null -> { customerDetails with CustomerName = None } 
    | name -> { customerDetails with CustomerName = Some name }

let getMonthDay (row:CustomerDetailsType.Row) (customerDetails: CustomerDetails) =
    match row.``Month Day`` with
    | day when day > 0.0 ->
        { customerDetails with Preference = Some (Day (int day)) }
    | _ -> customerDetails    

let getWeekDay (row:CustomerDetailsType.Row) (customerDetails: CustomerDetails) =
    let weekdayList = []
    
    let appendWeekday weekdayObj dayNum list =
        match weekdayObj with
        | null -> list
        | value when value.ToString().ToUpper().StartsWith "Y" -> dayNum::list
        | _ -> list
        
    let weekdays = weekdayList
                   |> appendWeekday row.Monday 1
                   |> appendWeekday row.Tuesday 2
                   |> appendWeekday row.Wednesday 3
                   |> appendWeekday row.Thursday 4
                   |> appendWeekday row.Friday 5
                   |> appendWeekday row.Saturday 6
                   |> appendWeekday row.Sunday 0

    match weekdays with
    | [] -> customerDetails
    | list -> { customerDetails with Preference = (Some (WeekDay list)) }

let getEveryDay (row:CustomerDetailsType.Row) (customerPreference: CustomerDetails) =
    match row.``Every Day`` with
    | null -> customerPreference
    | value when value.ToString().ToUpper().StartsWith "Y" -> { customerPreference with Preference = Some EveryDay }
    | _ -> customerPreference

let getNever (row:CustomerDetailsType.Row) (customerPreference: CustomerDetails) =
    match row.Never with
    | null -> customerPreference
    | value when value.ToString().ToUpper().StartsWith "Y" -> { customerPreference with Preference = Some Never }
    | _ -> customerPreference

// Assumes more frequent requests supercedes less frequent ones 
let readColumns (row:CustomerDetailsType.Row) (customerPreference: CustomerDetails) =
    customerPreference
    |> getCustomerName row
    |> getMonthDay row
    |> getWeekDay row
    |> getEveryDay row
    |> getNever row

let readRows (rows:CustomerDetailsType.Row array) =
    let rec readRowData index (list:CustomerDetails list) =
        if (index >= rows.Length) then
            list
        else
            let rowData = readColumns rows.[index] CustomerDetails.Default            
            readRowData (index + 1) (rowData::list)
            
    readRowData 0 []
    
let getCustomers (date:DateTime) (customerDetails:CustomerDetails) =    
    let weekDay =  int date.DayOfWeek
    let monthDay = date.Day
    
    match customerDetails with
    | { CustomerName = None } -> None
    | { Preference = Some Never } -> None
    | { Preference = Some EveryDay } -> Some customerDetails.CustomerName
    | { Preference = Some (WeekDay list) }
        when list |> List.contains weekDay -> Some customerDetails.CustomerName
    | { Preference = Some (Day day) }
        when day = monthDay -> Some customerDetails.CustomerName
    | _ -> None

let getCustomerOutput (startDate:DateTime) (customerData:CustomerDetails list) = seq { 
    let dates = seq { for i in 0..89 do yield startDate.AddDays(float i)}
    for date in dates do
        let customers = customerData                        
                        |> List.fold (fun customerList pref -> (pref |> getCustomers date) :: customerList) []
                        |> List.choose id
                        |> List.map (fun customer -> customer |> Option.get)
                        |> List.toArray
        let dateFormatted = date.ToString("ddd dd-MMMM-yyyy")
        yield dateFormatted, customers
}

let tryParseDate (dateStr:string) =
    let isOk, date = DateTime.TryParse(dateStr)
    match isOk with
    | true -> Some date
    | false -> None 

let displayWarning (message:string) =
    Console.ForegroundColor <- ConsoleColor.Yellow
    Console.WriteLine message    

let displayError (message:string) =
    Console.ForegroundColor <- ConsoleColor.Red
    Console.WriteLine message    

let getCustomersDetails =
    let dataSource = new CustomerDetailsType()
    let rows = dataSource.Data |> Seq.toArray
    rows |> readRows

let readDataAndDisplayReport startDate =
    let customerData = getCustomersDetails    
    for date, customers in getCustomerOutput startDate customerData do
        let customerList = String.Join (", ", customers)
        Console.WriteLine $"{date} - {customerList}"

let readDataAndSaveToFile startDate file =
    Console.WriteLine "\nStarted reading customer preference data"

    let customerData = getCustomersDetails    
    let outputData = seq {
        yield "Dates,CustomerNames"
        for date, customers in getCustomerOutput startDate customerData do
            let customerList = String.Join (",", customers |> Array.map (fun x -> $"\"{x}\""))
            yield $"\"{date}\",{customerList}"
    }
    
    try
        File.WriteAllLines (file, outputData)
        Console.WriteLine $"Customer notification details output to file '{file}'"
    with
        | :? Exception as ex -> displayError ex.Message    

let displayHelp () =
    Console.ForegroundColor <- ConsoleColor.Green
    Console.WriteLine "\n**********************************"
    Console.WriteLine "Usage:-"
    Console.WriteLine "Edit Excel File 'CustomerData.xlxs' with the customer preferences"
    Console.WriteLine "Then run program CustomerPreferenceApp.exe startdate [CsvFile]"
    Console.WriteLine "Parameter: StartDate - has format DD-MMM-YYYY"
    Console.WriteLine "Optional Parameter: CsvFile - output file name created to contain the results"
    Console.WriteLine "If no csvFile specified then results are output to console"
    Console.WriteLine "**********************************\n"
    Console.ForegroundColor <- ConsoleColor.White

[<EntryPoint>]
let main argv =        
    match argv.Length with
    | 0 -> displayHelp ()
    | 1 -> match tryParseDate argv.[0] with
           | Some date -> readDataAndDisplayReport date
           | None -> displayWarning $"Invalid date parameter {argv.[0]}"
    | 2 -> match tryParseDate argv.[0] with
           | Some date -> readDataAndSaveToFile date argv.[1]
           | None -> displayWarning $"Invalid date parameter {argv.[0]}"           
    | _ -> displayWarning "Invalid number of parameters"
    
    Console.ForegroundColor <- ConsoleColor.Green
    Console.WriteLine "\nProgram finished"
    Console.WriteLine "Press return to exit"
    Console.ReadKey() |> ignore
    0 // exit code