package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"io"
	"io/ioutil"
	"log"
	"os"
	"path"
	"strings"
	"unicode"
)

const (
	RAW_FILE          = "..\\登记表\\监考情况-RAW.xlsx"
	BBCC_PATH         = "..\\到场人员\\汇总"
	USER_NAME_SESSION = "B"
	CAN_NUM_SESSION   = "D"
	RES_SESSION       = "I"
	DEFAULT_SHEET     = "Sheet1"
	PHONE_SESSION     = "C"
	MATCCHED_XLSX     = "matched.xlsx"
)

type Student struct {
	userName string
	CanNum   string
	IsCan    string
	Phone    string
}

var students = make([]Student, 6000)
var logs *log.Logger //
var matched *excelize.File
var BBCCCanNum = make(map[string]string)

func main() {
	raw, nums := LoadRaw(RAW_FILE)
	samples := LoadBBC(BBCC_PATH)
	MatchSimples(raw, samples, nums)
	SaveMatched()
	Verify(samples, nums)
}

// 加载raw.xlsx
func LoadRaw(fileName string) (*excelize.File, int) {
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return f, 0
	}
	// 获取 Sheet1 上所有单元格
	rows, err := f.GetRows(DEFAULT_SHEET)
	lens := len(rows)
	for index := 2; index <= lens; index++ {
		if index == 0 || index == 1 {
			continue
		}
		name, err := f.GetCellValue(DEFAULT_SHEET, USER_NAME_SESSION+fmt.Sprintf("%d", index))
		can, err := f.GetCellValue(DEFAULT_SHEET, CAN_NUM_SESSION+fmt.Sprintf("%d", index))
		isCan, err := f.GetCellValue(DEFAULT_SHEET, RES_SESSION+fmt.Sprintf("%d", index))
		phone, err := f.GetCellValue(DEFAULT_SHEET, PHONE_SESSION+fmt.Sprintf("%d", index))
		if err != nil {
			fmt.Errorf(err.Error())
		}
		if can == "" {
			can = "U9999999"
		}
		stu := Student{
			userName: name,
			CanNum:   can[1:],
			IsCan:    isCan,
			Phone:    phone,
		}
		students[index] = stu
	}
	//for i := 2; i <= lens; i++ {
	//	fmt.Printf("index:%d,value:%s\t", i, students[i])
	//}
	return f, lens
}

// 加载path下的所有 bbcc 文件
func LoadBBC(path1 string) []string {
	// 获取所有文件名
	fileInfoList, err := ioutil.ReadDir(path1)
	if err != nil {
		fmt.Println(err.Error())
		return nil
	}

	var ret []string
	for _, file := range fileInfoList {
		str := strings.Split(file.Name(), "_")
		bbccCanName := str[1]
		f, err := excelize.OpenFile(path.Join(path1, file.Name()))
		if err != nil {
			fmt.Println(err)
			return nil
		}
		// 获取 Sheet1 上所有单元格
		rows, err := f.GetRows(DEFAULT_SHEET)
		lens := len(rows)
		for i := 1; i < lens; i++ {
			value, _ := f.GetCellValue(DEFAULT_SHEET, "A"+fmt.Sprintf("%d", i))
			if value == "" {
				continue
			}
			if len(value) == 0 {
				continue
			}
			if strings.Contains(value, "会议主题") {
				i = 7
				continue
			}
			ret = append(ret, value)
			BBCCCanNum[value] = bbccCanName
		}
	}
	// 排序
	//sort.Sort(sort.StringSlice(ret))
	for i, str := range ret {
		fmt.Println(i, "->", str)
	}
	return ret
}

// 匹配处理
func MatchSimples(raw *excelize.File, bbcc []string, nums int) {
	for index := 2; index < nums; index++ {
		if find(students[index], bbcc) == true {
			if students[index].IsCan == "" {
				raw.SetCellValue(DEFAULT_SHEET, RES_SESSION+fmt.Sprintf("%d", index), "有")
			} else {
				fmt.Printf("Already: %12s is %8s \n", students[index].userName, students[index].IsCan)
			}

		}
	}
	raw.SaveAs(".\\监考情况.xlsx")
}

// 查找
var mRow = 1

func find(stu Student, ref []string) bool {
	var err error
	var addr string
	for _, r := range ref {
		var sim float64
		SimilarText(r, stu.CanNum+stu.userName, &sim)
		if match(r, stu) == true {
			logs.Printf("matched: %v in %s with %f \n", stu, r, sim)
			data := []interface{}{
				stu.userName, stu.CanNum, stu.Phone, r, sim,
			}
			if addr, err = excelize.JoinCellName("A", mRow); err != nil {
				fmt.Println(err)
			}
			if err = matched.SetSheetRow(DEFAULT_SHEET, addr, &data); err != nil {
				fmt.Println(err)
			}
			mRow++
			//logs.Printf("matched: %v in %s with %f \n",students[index],bbcc)
			return true
		} else {
			//fmt.Println()
		}
	}
	return false
}

// 配对
func match(bbcc string, stu Student) bool {
	if strings.Index(bbcc, "****") == 3 && len(stu.Phone) > 4 && bbcc[:3] == stu.Phone[:3] && bbcc[7:] == stu.Phone[7:] { // 手机号判断,手机号对应就是
		//fmt.Printf("Phone: %12s -> %12s \n", bbcc, stu.Phone)
		return true
	}
	// 用户名对对应也是
	if strings.Contains(bbcc, stu.userName) {
		//fmt.Printf("Matched:%s\t%s\n",s1,d)
		return true
	}
	// 参会号策略多重判断
	if strings.Contains(bbcc, stu.CanNum) {
		//fmt.Printf("Matched:%s\t%s\n",s2,d)
		// 就是参会号
		if len(bbcc) == len("u"+stu.CanNum) {
			return true
		}
		var sim float64
		SimilarText(bbcc, stu.CanNum+stu.userName, &sim) // 判断参会号+用户名的相似度
		if sim > 66.0 {
			return true
		}
	}
	var simName, SimBoth float64
	SimilarText(bbcc, stu.CanNum+stu.userName, &SimBoth) // 判断参会号+用户名的相似度
	SimilarText(bbcc, stu.userName, &simName)            // 判断参会号+用户名的相似度
	if simName > 90.0 {
		fmt.Printf("Similar text: %s in %s \n", stu.userName, bbcc)
		return true
	}

	if SimBoth > 90.0 {
		fmt.Printf("Similar text: %s in %s \n", stu.CanNum+stu.userName, bbcc)
		return true
	}
	return false
}

// 计算文本相似度
func SimilarText(first, second string, percent *float64) int {
	var similarText func(string, string, int, int) int
	similarText = func(str1, str2 string, len1, len2 int) int {
		var sum, max int
		pos1, pos2 := 0, 0

		// Find the longest segment of the same section in two strings
		for i := 0; i < len1; i++ {
			for j := 0; j < len2; j++ {
				for l := 0; (i+l < len1) && (j+l < len2) && (str1[i+l] == str2[j+l]); l++ {
					if l+1 > max {
						max = l + 1
						pos1 = i
						pos2 = j
					}
				}
			}
		}

		if sum = max; sum > 0 {
			if pos1 > 0 && pos2 > 0 {
				sum += similarText(str1, str2, pos1, pos2)
			}
			if (pos1+max < len1) && (pos2+max < len2) {
				s1 := []byte(str1)
				s2 := []byte(str2)
				sum += similarText(string(s1[pos1+max:]), string(s2[pos2+max:]), len1-pos1-max, len2-pos2-max)
			}
		}

		return sum
	}

	l1, l2 := len(first), len(second)
	if l1+l2 == 0 {
		return 0
	}
	sim := similarText(first, second, l1, l2)
	if percent != nil {
		*percent = float64(sim*200) / float64(l1+l2)
	}
	return sim
}

func Verify(bbcc []string, num int) {
	var cnt = 1
	var maxSim = float64(0.0)
	var data []interface{}
	var err error
	var addr string
	file := excelize.NewFile()
	index := file.NewSheet(DEFAULT_SHEET)
	file.SetActiveSheet(index)
	for _, b := range bbcc {
		flg := false
		data = nil
		maxSim = 0.0
		for index := 2; index < num; index++ {
			var sim float64
			if IsFullChinese(b) == true {
				SimilarText(b, students[index].userName, &sim)
			} else {
				SimilarText(b, students[index].CanNum+students[index].userName, &sim)
			}
			if sim > maxSim {
				maxSim = sim
				data = []interface{}{
					b, BBCCCanNum[b], "|", students[index].userName, students[index].CanNum, students[index].Phone, sim,
				}
			}
			if match(b, students[index]) == true {
				flg = true
			}
		}
		if flg == false {
			//data = append(data, "0")
			logs.Printf("not verifyed: %s in %v \n", b, data)
			if addr, err = excelize.JoinCellName("A", cnt); err != nil {
				fmt.Println(err)
			}
			if err = file.SetSheetRow(DEFAULT_SHEET, addr, &data); err != nil {
				fmt.Println(err)
			}
			cnt++
		} else {
			//data = append(data, "1")
			//if addr, err = excelize.JoinCellName("A", cnt); err != nil {
			//	fmt.Println(err)
			//}
			//if err = file.SetSheetRow(DEFAULT_SHEET, addr, &data); err != nil {
			//	fmt.Println(err)
			//}
			//cnt++
		}
	}
	file.SaveAs("NoMatchInBBCC.xlsx")
}

func init() {
	file, err := os.OpenFile("logs.txt",
		os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		log.Fatalln("Failed to open error log file:", err)
	}
	logs = log.New(io.MultiWriter(file, os.Stdout),
		"", 0)

	matched = excelize.NewFile()
	index := matched.NewSheet(DEFAULT_SHEET)
	matched.SetActiveSheet(index)
}

func SaveMatched() {
	err := matched.SaveAs(MATCCHED_XLSX)
	if err != nil {
		fmt.Println(err.Error())
	}
}

func IsFullChinese(str string) bool {
	var count int
	var lenth int
	for _, v := range str {
		lenth++
		if unicode.Is(unicode.Han, v) {
			count++
		}
	}
	return count == lenth
}
