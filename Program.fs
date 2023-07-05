open System.IO
open System.Net.Http
open MailKit
open MailKit.Net.Imap
open System

let run () =
    use client = new ImapClient()
    
    // Output is in german as target audience speaks german and we have no localization
    Console.Write "Server Adresse: "
    let address = Console.ReadLine()
   
    Console.Write "Server Port: "
    let port = Console.ReadLine() |> int
   
    client.Connect(address, port)
    
    Console.Write "Nutzername: "
    let userName = Console.ReadLine()
    
    Console.Write "Passwort: "
    let password = Console.ReadLine()

    printfn "Anmeldung..."
    client.Authenticate(userName, password)
    printfn "Erfolgreich angemeldet"
    
    printfn "Öffne Eingang..."
    let inbox = client.Inbox

    inbox.Open FolderAccess.ReadOnly |> ignore
    printfn "Eingang geöffnet"
    
    printfn "Lade neusten 5 Emails..."
    for index = inbox.Count - 1 downto max (inbox.Count - 5) 0 do
        let message = inbox.GetMessage(index)
        let path = Path.GetFullPath $"./{message.Subject}.eml"
        message.WriteTo path

        printfn $"""Email gespeichert unter "{path}" """
    
    printfn "Programm beendet"
        
run ()