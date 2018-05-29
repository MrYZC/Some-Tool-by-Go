package main

import (
	"bufio"
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"io/ioutil"
	"math"
	"os"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var binDataSlice []uint16
var tansferRow float64
var tansferAllRow float64

var numRow int
var bufferColA, bufferColB, bufferColC, bufferColD, bufferColE, bufferColF, bufferColG, bufferColH, bufferColI, bufferColJ, bufferColK bytes.Buffer
var ColA, ColB, ColC, ColD, ColE, ColF, ColG, ColH, ColI, ColJ, ColK string
var ListBuf uint16
var sliceByte = make([]byte, 2)
var pathBin string = "./airData1.bin"
var pathTxt string = "./airData1.txt"
var pathRecord string = "./上传日志.txt"

/*
***author:  apple********
***todo:判断文件是否存在***
 */
func PathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}

func createXlsx(numRowRead int, arrRev *[]uint16) {
	PathIsEis, _ := PathExists("./空气检测.xlsx")
	if PathIsEis == false {
		xlsx := excelize.NewFile()
		//xlsx.NewSheet("空气检测.Excel")
		xlsx.SetCellValue("Sheet1", "A1", "年")
		xlsx.SetCellValue("Sheet1", "B1", "月/日")
		xlsx.SetCellValue("Sheet1", "C1", "时/分")
		xlsx.SetCellValue("Sheet1", "D1", "处理前：")
		xlsx.SetCellValue("Sheet1", "E1", "苯浓度")
		xlsx.SetCellValue("Sheet1", "F1", "CO2浓度")
		xlsx.SetCellValue("Sheet1", "G1", "TOVC浓度")
		xlsx.SetCellValue("Sheet1", "H1", "处理后:")
		xlsx.SetCellValue("Sheet1", "I1", "苯浓度")
		xlsx.SetCellValue("Sheet1", "J1", "CO2浓度")
		xlsx.SetCellValue("Sheet1", "K1", "TVOC浓度")
		for ii := 2; ii < numRow+2; ii++ {
			strii := strconv.Itoa(ii)
			bufferColA.WriteString("A")
			bufferColA.WriteString(strii)
			bufferColB.WriteString("B")
			bufferColB.WriteString(strii)
			bufferColC.WriteString("C")
			bufferColC.WriteString(strii)
			bufferColD.WriteString("D")
			bufferColD.WriteString(strii)
			bufferColE.WriteString("E")
			bufferColE.WriteString(strii)
			bufferColF.WriteString("F")
			bufferColF.WriteString(strii)
			bufferColG.WriteString("G")
			bufferColG.WriteString(strii)
			bufferColH.WriteString("H")
			bufferColH.WriteString(strii)
			bufferColI.WriteString("I")
			bufferColI.WriteString(strii)
			bufferColJ.WriteString("J")
			bufferColJ.WriteString(strii)
			bufferColK.WriteString("K")
			bufferColK.WriteString(strii)
			ColA = bufferColA.String()
			ColB = bufferColB.String()
			ColC = bufferColC.String()
			ColD = bufferColD.String()
			ColE = bufferColE.String()
			ColF = bufferColF.String()
			ColG = bufferColG.String()
			ColH = bufferColH.String()
			ColI = bufferColI.String()
			ColJ = bufferColJ.String()
			ColK = bufferColK.String()
			xlsx.SetCellValue("Sheet1", ColA, (*arrRev)[(ii-2)*9])
			xlsx.SetCellValue("Sheet1", ColB, (*arrRev)[(ii-2)*9+1])
			xlsx.SetCellValue("Sheet1", ColC, (*arrRev)[(ii-2)*9+2])
			xlsx.SetCellValue("Sheet1", ColD, "")
			xlsx.SetCellValue("Sheet1", ColE, (*arrRev)[(ii-2)*9+3])
			xlsx.SetCellValue("Sheet1", ColF, (*arrRev)[(ii-2)*9+4])
			xlsx.SetCellValue("Sheet1", ColG, (*arrRev)[(ii-2)*9+5])
			xlsx.SetCellValue("Sheet1", ColH, "")
			xlsx.SetCellValue("Sheet1", ColI, (*arrRev)[(ii-2)*9+6])
			xlsx.SetCellValue("Sheet1", ColJ, (*arrRev)[(ii-2)*9+7])
			xlsx.SetCellValue("Sheet1", ColK, (*arrRev)[(ii-2)*9+8])
			bufferColA.Reset()
			bufferColB.Reset()
			bufferColC.Reset()
			bufferColD.Reset()
			bufferColE.Reset()
			bufferColF.Reset()
			bufferColG.Reset()
			bufferColH.Reset()
			bufferColI.Reset()
			bufferColJ.Reset()
			bufferColK.Reset()
			fmt.Printf("生成Excel进度: %d%%\n", 100*ii/(numRow+1))
		}
		err := xlsx.SaveAs("./空气检测.xlsx")
		if err != nil {
			fmt.Println("Save xlsx:", err)
		}
	} else {
		xlsx, err := excelize.OpenFile("./空气检测.xlsx")
		if err != nil {
			fmt.Println("Open File falied:", err)
		}
		rowLast := xlsx.GetRows("Sheet1")
		RowLastNum := len(rowLast)
		if err != nil {
			fmt.Println(err)
			return
		}
		for ii := RowLastNum; ii < numRow+2; ii++ {
			strii := strconv.Itoa(ii)
			bufferColA.WriteString("A")
			bufferColA.WriteString(strii)
			bufferColB.WriteString("B")
			bufferColB.WriteString(strii)
			bufferColC.WriteString("C")
			bufferColC.WriteString(strii)
			bufferColD.WriteString("D")
			bufferColD.WriteString(strii)
			bufferColE.WriteString("E")
			bufferColE.WriteString(strii)
			bufferColF.WriteString("F")
			bufferColF.WriteString(strii)
			bufferColG.WriteString("G")
			bufferColG.WriteString(strii)
			bufferColH.WriteString("H")
			bufferColH.WriteString(strii)
			bufferColI.WriteString("I")
			bufferColI.WriteString(strii)
			bufferColJ.WriteString("J")
			bufferColJ.WriteString(strii)
			bufferColK.WriteString("K")
			bufferColK.WriteString(strii)
			ColA = bufferColA.String()
			ColB = bufferColB.String()
			ColC = bufferColC.String()
			ColD = bufferColD.String()
			ColE = bufferColE.String()
			ColF = bufferColF.String()
			ColG = bufferColG.String()
			ColH = bufferColH.String()
			ColI = bufferColI.String()
			ColJ = bufferColJ.String()
			ColK = bufferColK.String()
			xlsx.SetCellValue("Sheet1", ColA, (*arrRev)[(ii-2)*9])
			xlsx.SetCellValue("Sheet1", ColB, (*arrRev)[(ii-2)*9+1])
			xlsx.SetCellValue("Sheet1", ColC, (*arrRev)[(ii-2)*9+2])
			xlsx.SetCellValue("Sheet1", ColD, "")
			xlsx.SetCellValue("Sheet1", ColE, (*arrRev)[(ii-2)*9+3])
			xlsx.SetCellValue("Sheet1", ColF, (*arrRev)[(ii-2)*9+4])
			xlsx.SetCellValue("Sheet1", ColG, (*arrRev)[(ii-2)*9+5])
			xlsx.SetCellValue("Sheet1", ColH, "")
			xlsx.SetCellValue("Sheet1", ColI, (*arrRev)[(ii-2)*9+6])
			xlsx.SetCellValue("Sheet1", ColJ, (*arrRev)[(ii-2)*9+7])
			xlsx.SetCellValue("Sheet1", ColK, (*arrRev)[(ii-2)*9+8])
			bufferColA.Reset()
			bufferColB.Reset()
			bufferColC.Reset()
			bufferColD.Reset()
			bufferColE.Reset()
			bufferColF.Reset()
			bufferColG.Reset()
			bufferColH.Reset()
			bufferColI.Reset()
			bufferColJ.Reset()
			bufferColK.Reset()
			fmt.Printf("生成Excel进度: %d%%\n", 100*ii/(numRow+1))
		}
		err1 := xlsx.SaveAs("./空气检测.xlsx")
		if err1 != nil {
			fmt.Println(err1)
		}
	}
}

/*
****author: apple ***************
****todo: ascii将int型[]byte转化为int**
 */
func asciiByteToInt(b []byte) []int {
	var ascB []int
	for indexSlice := range b {
		switch b[indexSlice] {
		case 48:
			ascB = append(ascB, 0)
		case 49:
			ascB = append(ascB, 1)
		case 50:
			ascB = append(ascB, 2)
		case 51:
			ascB = append(ascB, 3)
		case 52:
			ascB = append(ascB, 4)
		case 53:
			ascB = append(ascB, 5)
		case 54:
			ascB = append(ascB, 6)
		case 55:
			ascB = append(ascB, 7)
		case 56:
			ascB = append(ascB, 8)
		case 57:
			ascB = append(ascB, 9)
		}
	}
	return ascB
}

/*
****author: apple ***************
****todo: ascii将int型[]byte转化为int**
 */

func intSliceToInt(intSlice []int) int {
	sum := 0
	for intSliceIndex := range intSlice {
		sum = sum + intSlice[intSliceIndex]*int(math.Pow(10, float64(len(intSlice)-intSliceIndex-1)))
	}
	return sum
}

/***author: apple ***************
***todo: record the Download record*****
 */
func recordDownLoad() int {
	_, err := os.OpenFile(pathRecord, os.O_RDWR|os.O_CREATE, 0666)
	if err != nil {
		fmt.Println("create or open：", err)
	}
	bufferFileRecord, err := ioutil.ReadFile(pathRecord)
	if err != nil {
		fmt.Println("read ioutil:", err)
	}
	if bufferFileRecord == nil {
		return 0
	} else {
		for tt := range bufferFileRecord {
			if bufferFileRecord[tt] == 10 && tt >= (len(bufferFileRecord)-11) {
				allNumBitSlice := bufferFileRecord[tt+1:]
				fmt.Println(intSliceToInt(asciiByteToInt(allNumBitSlice)))
				return intSliceToInt(asciiByteToInt(allNumBitSlice))
			}
		}
		return 0
	}
}

/*
***author: apple ***************
***todo: record the Upload record *****
 */
var nextLine = []byte{}

func recordUpload(numRowUpLoad int, arrRev *[]uint16) {
	fileRecord, err := os.OpenFile(pathRecord, os.O_RDWR|os.O_CREATE|os.O_APPEND, 0644)
	if err != nil {
		fmt.Println("Create or Open fileRecord:", err)
	}
	defer fileRecord.Close()
	nextLine = append(nextLine, 0x0a)
	fileRecord.Write(nextLine)
	fileRecord.WriteString("上传日志的行数：")
	fileRecord.Write(nextLine)
	fileRecord.WriteString(strconv.Itoa(numRowUpLoad + 1))
}

func main() {
	logLine := recordDownLoad() // 日志记录上次所读的行
	fmt.Println(logLine)
	fileBin, err := os.OpenFile(pathBin, os.O_RDWR, 0666)
	if err != nil {
		fmt.Println("解析bin文件失败, 请查看是否把单片机的UserDb.bin文件拷贝到当前文件夹:", err)
		return
	}
	fileTxt, err := os.OpenFile(pathTxt, os.O_RDWR|os.O_CREATE, 0666)
	defer fileBin.Close()
	defer fileTxt.Close()
	if err != nil {
		fmt.Println("openFile or  createFile failed:", err)
	}
	bufferDataBin := bufio.NewReader(fileBin)
	for {
		bufList1, err1 := bufferDataBin.ReadByte()
		if err1 != nil && err1 != io.EOF {
			fmt.Println("read bufData err:", err1)
			break
		}
		if err1 == io.EOF {
			fmt.Println("读取传感器数据完毕...  请等3秒")
			time.Sleep(time.Second)
			break
		}
		bufList2, err2 := bufferDataBin.ReadByte()
		if err2 != nil && err2 != io.EOF {
			fmt.Println("read bufData err:", err2)
			break
		}
		if err2 == io.EOF {
			fmt.Println("读取传感器数据完...   请等待3秒")
			time.Sleep(3 * time.Second)
			break
		}
		sliceByte[0] = bufList1
		sliceByte[1] = bufList2
		ListBuf = uint16(binary.BigEndian.Uint16(sliceByte))
		if err != nil {
			fmt.Println("binaryRead err:", err)
			continue
		}
		binDataSlice = append(binDataSlice, ListBuf)
	}
	fmt.Println("len:", len(binDataSlice))
	fmt.Println(logLine)
	numRowAll := int(math.Floor(float64(len(binDataSlice) / 9)))
	binDataSlice = binDataSlice[9*logLine:]
	tansferAllRow = math.Floor(float64(len(binDataSlice) / 9))
	for ii := range binDataSlice {
		if ii%9 != 8 {
			fmt.Fprintf(fileTxt, "%6d\t", binDataSlice[ii])
		}
		if ii%9 == 8 {
			tansferRow = float64((ii + 1) / 9)
			fmt.Fprintf(fileTxt, "%6d\n", binDataSlice[ii])
			fmt.Printf("转换第%d行数据成TXT文本完毕, 已完成%f%%\n", (ii+1)/9+9*logLine, 100*tansferRow/tansferAllRow)
			numRow++
		}
		if ii == len(binDataSlice)-1 {
			fmt.Println("所有数据已转换成TXT文本完毕。 请等待3秒准备把数据写入Excel表格")
		}
	}
	//time.Sleep(3 * time.Second)
	createXlsx(numRow, &binDataSlice)
	recordUpload(numRowAll, &binDataSlice)
}
