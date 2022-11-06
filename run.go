package main

import (
    "fmt"
    "bufio"
    "log"
    "os"
    "time"
    "io/ioutil"
    "strings"
    "strconv"
    
    cp "github.com/otiai10/copy"
    "github.com/xuri/excelize/v2"
    "github.com/go-ping/ping"
)

var ping_total = 0

func writeData(excel_name string, sheet_name string, map_1 map[string]string, map_2 map[string]string, 
                    map_3 map[string]string, map_4 map[string]string) {
    f, err := excelize.OpenFile(excel_name)
    if err != nil {
        fmt.Println(err)
        return
    }
    defer func() {
        // Close the spreadsheet.
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()

    // Get all the rows in the sheet.
    rows, err := f.GetRows(sheet_name)
    if err != nil {
        fmt.Println(err)
        return
    }

    length := len(rows)
    for i:=1; i <= length; i+=1 {
        tmp := "D" + strconv.Itoa(i)
        cell, err := f.GetCellValue(sheet_name, tmp)
        if err != nil {
            fmt.Println(err)
            return
        }
        // fmt.Println(cell)
        val1, ok1 := map_1[cell]
        if ok1 {
            tmp := "G" + strconv.Itoa(i)
            f.SetCellValue(sheet_name, tmp, val1)
        }

        val2, ok2 := map_2[cell]
        if ok2 {
            tmp := "H" + strconv.Itoa(i)
            f.SetCellValue(sheet_name, tmp, val2)
        }

        val3, ok3 := map_3[cell]
        if ok3 {
            tmp := "J" + strconv.Itoa(i)
            f.SetCellValue(sheet_name, tmp, val3)
        }

        val4, ok4 := map_4[cell]
        if ok4 {
            tmp := "I" + strconv.Itoa(i)
            f.SetCellValue(sheet_name, tmp, val4)
        }
    }

    if err := f.SaveAs("Updated.xlsx"); err != nil {
        fmt.Println(err)
    }

}

func IOReadLog(root string) (map[string]string, map[string]string, error) {
    fileInfo, err := ioutil.ReadDir(root)
    if err != nil {
        log.Fatal(err)
    }

    var map_1 = make(map[string]string)
    var map_2 = make(map[string]string)
    for _, file := range fileInfo {
        // files = append(files, file.Name())
        f, err := os.Open(root + `\` + file.Name())
        if err != nil {
            log.Fatal(err)
        }

        defer f.Close()
        scanner := bufio.NewScanner(f)
        for scanner.Scan() {
            // fmt.Println(scanner.Text())
            array := strings.Fields(scanner.Text())
            if len(array) == 5 {
                if array[1] == `heropos` {
                    map_1[array[0]] = array[2] + ` ` + array[3] + ` ` + array[4]
                } else if array[1] == `heropos2` {
                    map_2[array[0]] = array[2] + ` ` + array[3] + ` ` + array[4]
                }
            }
            // fmt.Println(array, len(array))
        }
        if err := scanner.Err(); err != nil {
            log.Fatal(err)
        }
    }
    return map_1, map_2, err
}

func IOReadReg(root string) (map[string]string, map[string]string, error) {
    fileInfo, err := ioutil.ReadDir(root)
    if err != nil {
        log.Fatal(err)
    }

    var map_1 = make(map[string]string)
    var map_2 = make(map[string]string)
    for _, file := range fileInfo {
        // files = append(files, file.Name())
        f, err := os.Open(root + `\` + file.Name())
        if err != nil {
            log.Fatal(err)
        }

        defer f.Close()
        scanner := bufio.NewScanner(f)
        for scanner.Scan() {
            // fmt.Println(scanner.Text())
            array := strings.Fields(scanner.Text())
            if len(array) == 5 {
                map_1[array[0]] = array[2] + ` ` + array[3] + ` ` + array[4]
                map_2[array[0]] = array[1]
            }
            // fmt.Println(array, len(array))
    	}
        if err := scanner.Err(); err != nil {
            log.Fatal(err)
        }
    }
    return map_1, map_2, err
}



func getPingStat(host_name string, c chan string) {
    var text string
    pinger, err := ping.NewPinger(host_name)
    if err != nil {
        // fmt.Println("ERROR:", err)
        text = host_name + "@ERROR:" + err.Error() + "\n"
        c <- text
        return
    }

    pinger.Count = 5
    pinger.SetPrivileged(true)
    pinger.Timeout = time.Second*5
    err = pinger.Run()
    if err != nil {
        fmt.Println("Failed to ping target host:", err)
    }
    pinger.Stop()
    stats := pinger.Statistics()
    text = stats.Addr + "@" +strconv.Itoa(stats.PacketsSent) + " packets transmitted, " + 
                strconv.Itoa(stats.PacketsRecv) + " packets received, " + 
                strconv.FormatFloat(stats.PacketLoss, 'g', 3, 64) + "% packet loss, IP:" + 
                pinger.IPAddr().String() + "\n"

    c <- text
}


func pingBySheet(excel_name string, sheet_name string, result chan string) {
    f, err := excelize.OpenFile(excel_name)
    if err != nil {
        fmt.Println(err)
        return
    }

    defer func() {
        // Close the spreadsheet.
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()

    rows, err := f.GetRows(sheet_name)
    if err != nil {
        fmt.Println(err)
        return
    }
    length := len(rows)
    ping_total += length
    
    for i:=1; i <= length; i+=1 {
        tmp := "D" + strconv.Itoa(i)
        cell, err := f.GetCellValue(sheet_name, tmp)
        if err != nil {
            fmt.Println(err)
            return
        }
        go getPingStat(cell, result)
    }

}

func readPingStat(ping_file string)(map[string]string) {

    var map_stat = make(map[string]string)

    f, err := os.Open(ping_file)
    if err != nil {
        log.Fatal(err)
    }
    defer f.Close()

    scanner := bufio.NewScanner(f)
    for scanner.Scan() {
        array := strings.Split(scanner.Text(), "@")
        if len(array) == 2 {
            map_stat[array[0]] = array[1]
        }
    }
    if err := scanner.Err(); err != nil {
        log.Fatal(err)
    }

    return map_stat
}

func writePingStat(excel_name string, sheet_name string, map_stat map[string]string) {
    f, err := excelize.OpenFile(excel_name)
    if err != nil {
        fmt.Println(err)
        return
    }
    defer func() {
        // Close the spreadsheet.
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()

    // Get all the rows in the sheet.
    rows, err := f.GetRows(sheet_name)
    if err != nil {
        fmt.Println(err)
        return
    }

    length := len(rows)
    for i:=1; i <= length; i+=1 {
        tmp := "D" + strconv.Itoa(i)
        cell, err := f.GetCellValue(sheet_name, tmp)
        if err != nil {
            fmt.Println(err)
            return
        }
        val, ok := map_stat[cell]
        if ok {
            tmp := "F" + strconv.Itoa(i)
            f.SetCellValue(sheet_name, tmp, val)
        }
    }

    if err := f.SaveAs("Updated.xlsx"); err != nil {
        fmt.Println(err)
    }
}


func main() {
    var (
        err error
        map_1 map[string]string
        map_2 map[string]string
        map_3 map[string]string
        map_4 map[string]string
    )

    log_path_sc := `//10.102.241.37/Scripts/LastLogin_Logs`
    log_path_cod := `//10.102.241.48/Scripts/LastLogin_Logs`
    log_path_alt := `//10.102.241.83/Scripts/LastLogin_Logs`
    reg_path_sc := `//10.102.241.37/Scripts/Reg_Logs`
    reg_path_cod := `//10.102.241.48/Scripts/Reg_Logs`
    reg_path_alt := `//10.102.241.83/Scripts/Reg_Logs`
    log_des_sc := `.\Logs_SC`
    log_des_cod := `.\Logs_COD`
    log_des_alt := `.\Logs_ALT`
    reg_des_sc := `.\Reg_SC`
    reg_des_cod := `.\Reg_COD`
    reg_des_alt := `.\Reg_ALT`

    //SC
    fmt.Print(`Copying SC login logs ...`)
    err = cp.Copy(log_path_sc, log_des_sc)
    if err != nil {
        panic(err)
    }
    fmt.Println(`Done`)
    fmt.Print(`Copying SC registry logs ...`)
    err = cp.Copy(reg_path_sc, reg_des_sc)
    if err != nil {
        panic(err)
    }
    fmt.Println(`Done`)
    fmt.Print(`Updating SC to excel ...`)
    map_1, map_2, err = IOReadLog(log_des_sc)
    if err != nil {
        panic(err)
    }
    map_3, map_4, err = IOReadReg(reg_des_sc)
    if err != nil {
        panic(err)
    }
    writeData(`POS Workstations.xlsx`, `MSC`, map_1, map_2, map_3, map_4)
    fmt.Println(`Done`)

    //COD
    fmt.Print(`Copying COD login logs ...`)
    err = cp.Copy(log_path_cod, log_des_cod)
    if err != nil {
        panic(err)
    }
    fmt.Println(`Done`)
    fmt.Print(`Copying COD registry logs ...`)
    err = cp.Copy(reg_path_cod, reg_des_cod)
    if err != nil {
        panic(err)
    }
    fmt.Println(`Done`)
    fmt.Print(`Updating COD to excel ...`)
    map_1, map_2, err = IOReadLog(log_des_cod)
    if err != nil {
        panic(err)
    }
    map_3, map_4, err = IOReadReg(reg_des_cod)
    if err != nil {
        panic(err)
    }
    writeData(`Updated.xlsx`, `COD`, map_1, map_2, map_3, map_4)
    fmt.Println(`Done`)

    //ALT
    fmt.Print(`Copying Altira & Mocha login logs ...`)
    err = cp.Copy(log_path_alt, log_des_alt)
    if err != nil {
        panic(err)
    }
    fmt.Println(`Done`)
    fmt.Print(`Copying Altira & Mocha registry logs ...`)
    err = cp.Copy(reg_path_alt, reg_des_alt)
    if err != nil {
        panic(err)
    }
    fmt.Println(`Done`)
    fmt.Print(`Updating Altira & Mocha to excel ...`)
    map_1, map_2, err = IOReadLog(log_des_alt)
    if err != nil {
        panic(err)
    }
    map_3, map_4, err = IOReadReg(reg_des_alt)
    if err != nil {
        panic(err)
    }
    writeData(`Updated.xlsx`, `Altira`, map_1, map_2, map_3, map_4)
    fmt.Println(`Done`)
    
    //channel
    var result = make(chan string)

    //Get Ping Status
    fmt.Print(`Concurrently getting Ping status...`)
    pingBySheet(`Updated.xlsx`, `MSC`, result)
    pingBySheet(`Updated.xlsx`, `COD`, result)
    pingBySheet(`Updated.xlsx`, `Altira`, result)

    f_txt, err := os.OpenFile(`pingStat.txt`, os.O_WRONLY|os.O_CREATE, 0644)//os.O_APPEND
    if err != nil {
        panic(err)
    }
    defer f_txt.Close()

    var i = 1
    for s := range result {
        // println(s)
        if _, err = f_txt.WriteString(s); err != nil {
            panic(err)
        }

        if i >= ping_total {
            close(result)
        }
        i++
    }
    fmt.Println(`Done`)

    //Write Ping Status
    fmt.Print(`Writing Ping Status to  excel ...`)
    map_stat := readPingStat(`pingStat.txt`)
    writePingStat(`Updated.xlsx`, `MSC`, map_stat)
    writePingStat(`Updated.xlsx`, `COD`, map_stat)
    writePingStat(`Updated.xlsx`, `Altira`, map_stat)
    fmt.Println(`Done`)

    fmt.Println(`Finish`)
    time.Sleep(5 * time.Second)

}