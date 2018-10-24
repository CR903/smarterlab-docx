package docx

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"strconv"
	"strings"
)

type ZipData interface {
	files() []*zip.File
	close() error
}

//Type for in memory zip files
type ZipInMemory struct {
	data *zip.Reader
}

func (d ZipInMemory) files() []*zip.File {
	return d.data.File
}

//Since there is nothing to close for in memory, just nil the data and return nil
func (d ZipInMemory) close() error {
	d.data = nil
	return nil
}

//Type for zip files read from disk
type ZipFile struct {
	data *zip.ReadCloser
}

func (d ZipFile) files() []*zip.File {
	return d.data.File
}

func (d ZipFile) close() error {
	return d.data.Close()
}

type ReplaceDocx struct {
	ZipReader ZipData
	Content   string
}

type Text struct {
	Words    string `json:"word"`
	Color    string `json:"color"`
	Size     string `json:"size"`
	Isbold   bool   `json:"isbold"`
	IsCenter bool   `json:"iscenter"`
}

type TableTHead struct {
	TData interface{} `json:"tdata"`
	TDW   int         `json:"tdw"`
}

//TableTD descripes every block of the table
type TableTD struct {
	//TData refers block's element
	TData []interface{} `json:"tdata"`
	//TDBG refers block's background
	TDBG int `json:"tdbg"`
	TDW  int `json:"tdw"`
	TDM  int `json:"tdm"` // 0 - 无 1 - 开始 2 - 结束
}

//Table include table configuration.
type Table struct {
	//Tbname  is the name of the table
	Tbname string `json:"tbname"`
	//Text OR Image in the sanme line
	Inline bool `json:"inline"`
	//Table data except table head
	TableBody [][]*TableTD `json:"tablebody"`
	//Table head data
	TableHead [][]*TableTHead `json:"tablehead"`
	//Thcenter set table head center word
	Thcenter bool `json:"thcenter"`
}

func (r *ReplaceDocx) Editable() *Docx {
	return &Docx{
		Files:   r.ZipReader.files(),
		Content: r.Content,
	}
}

func (r *ReplaceDocx) Close() error {
	return r.ZipReader.close()
}

type Docx struct {
	Files   []*zip.File
	Content string
}

func ReadDocxFromMemory(data io.ReaderAt, size int64) (*ReplaceDocx, error) {
	reader, err := zip.NewReader(data, size)
	if err != nil {
		return nil, err
	}
	zipData := ZipInMemory{data: reader}
	return ReadDocx(zipData)
}

func ReadDocxFile(path string) (*ReplaceDocx, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	zipData := ZipFile{data: reader}
	return ReadDocx(zipData)
}

func ReadDocx(reader ZipData) (*ReplaceDocx, error) {
	content, err := readText(reader.files())
	if err != nil {
		return nil, err
	}

	return &ReplaceDocx{ZipReader: reader, Content: content}, nil
}

func readText(files []*zip.File) (text string, err error) {
	var documentFile *zip.File
	documentFile, err = retrieveWordDoc(files)
	if err != nil {
		return text, err
	}
	var documentReader io.ReadCloser
	documentReader, err = documentFile.Open()
	if err != nil {
		return text, err
	}

	text, err = wordDocToString(documentReader)
	return
}

func wordDocToString(reader io.Reader) (string, error) {
	b, err := ioutil.ReadAll(reader)
	if err != nil {
		return "", err
	}
	return string(b), nil
}

func retrieveWordDoc(files []*zip.File) (file *zip.File, err error) {
	for _, f := range files {
		if f.Name == "word/document.xml" {
			file = f
		}
	}
	if file == nil {
		err = errors.New("document.xml file not found")
	}
	return
}

func (d *Docx) ReplaceRaw(oldString string, newString string, num int) {
	d.Content = strings.Replace(d.Content, fmt.Sprintf("{{%s}}", oldString), newString, num)
}

func (d *Docx) Replace(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}
	d.Content = strings.Replace(d.Content, fmt.Sprintf("{{%s}}", oldString), newString, num)

	return nil
}

func (d *Docx) ReplaceXML(oldString string, newString string, num int) {
	d.Content = strings.Replace(d.Content, fmt.Sprintf("{{%s}}", oldString), newString, num)
}

//WriteTable  ==表格的格式
func (d *Docx) ReplaceTable(table *Table) error {
	XMLTable := bytes.Buffer{}
	inline := table.Inline
	tableBody := table.TableBody
	tableHead := table.TableHead
	var used bool
	used = false

	//handle TableHead :Split with TableBody
	if tableHead != nil {
		XMLTable.WriteString(XMLTableHead)
		XMLTable.WriteString(XMLTableGridBegin)
		for i, h := range tableHead {
			gcw := fmt.Sprintf(XMLTableGridCol, strconv.FormatInt(int64(h[i].TDW), 10))
			XMLTable.WriteString(gcw)
		}
		XMLTable.WriteString(XMLTableGridEnd)

		XMLTable.WriteString(XMLTableHeadTR)
		for j, rowdata := range tableHead {
			thw := fmt.Sprintf(XMLHeadTableTDBegin, strconv.FormatInt(int64(rowdata[j].TDW), 10))
			XMLTable.WriteString(thw)
			if inline {
				if table.Thcenter {
					XMLTable.WriteString(XMLHeadTableTDBegin2C)
				} else {
					XMLTable.WriteString(XMLHeadTableTDBegin2)
				}
			}
			for _, rowEle := range rowdata {
				if !inline {
					if table.Thcenter {
						XMLTable.WriteString(XMLHeadTableTDBegin2C)
					} else {
						XMLTable.WriteString(XMLHeadTableTDBegin2)
					}
				}
				if text, ok := rowEle.TData.(*Text); ok {
					//not
					color := text.Color
					size := text.Size
					word := text.Words
					var data string
					if text.IsCenter {
						if text.Isbold {
							data = fmt.Sprintf(XMLHeadtableTDTextBC, color, size, size, word)
						} else {
							data = fmt.Sprintf(XMLHeadtableTDTextC, color, size, size, word)
						}
					} else {
						if text.Isbold {
							data = fmt.Sprintf(XMLHeadtableTDTextB, color, size, size, word)
						} else {
							data = fmt.Sprintf(XMLHeadtableTDText, color, size, size, word)
						}
					}
					XMLTable.WriteString(data)
				}
				if !inline {
					XMLTable.WriteString(XMLIMGtail)
				}
			}
			if inline {
				XMLTable.WriteString(XMLIMGtail)
			}
			XMLTable.WriteString(XMLHeadTableTDEnd)
		}
		XMLTable.WriteString(XMLTableEndTR)
	} else {
		XMLTable.WriteString(XMLTableNoHead)
		XMLTable.WriteString(XMLTableGridBegin)
		if len(tableBody) > 0 {
			for _, tb := range tableBody[0] {
				gcw := fmt.Sprintf(XMLTableGridCol, strconv.Itoa(tb.TDW))
				XMLTable.WriteString(gcw)
			}
			XMLTable.WriteString(XMLTableGridEnd)
		}
	}
	//Generate formation
	for _, v := range tableBody {
		XMLTable.WriteString(XMLTableTR)
		for _, vv := range v {
			//td bg
			var td string
			if vv.TDM == 1 {
				td = fmt.Sprintf(XMLTableMergeSTD, strconv.FormatInt(int64(vv.TDW), 10), strconv.FormatInt(int64(vv.TDBG), 10))
			} else if vv.TDM == 2 {
				td = fmt.Sprintf(XMLTableMergeCTD, strconv.FormatInt(int64(vv.TDW), 10), strconv.FormatInt(int64(vv.TDBG), 10))
			} else {
				td = fmt.Sprintf(XMLTableTD, strconv.FormatInt(int64(vv.TDW), 10), strconv.FormatInt(int64(vv.TDBG), 10))
			}
			XMLTable.WriteString(td)
			if vv.TDM < 2 {
				tds := 0
				for _, vvv := range vv.TData {
					table, ok := vvv.(*Table)
					if !inline && !ok {
						XMLTable.WriteString(XMLTableTD2)
					}
					if inline && !ok && tds == 0 {
						XMLTable.WriteString(XMLTableTD2)
					}
					//if td is a table
					if ok {
						//end with table
						used = true
						tablestr, err := replaceTableToBuffer(table)
						if err != nil {
							return err
						}
						XMLTable.WriteString(tablestr)
						// FIXME: magic operation
						XMLTable.WriteString(XMLMagicFooter)
						//image or text
					} else {
						if text, ko := vvv.(*Text); ko {
							if text.IsCenter {
								if text.Isbold {
									XMLTable.WriteString(XMLHeadtableTDTextBC)
								} else {
									XMLTable.WriteString(XMLHeadtableTDTextC)
								}
							} else {
								if text.Isbold {
									XMLTable.WriteString(XMLHeadtableTDTextB)
								} else {
									XMLTable.WriteString(XMLHeadtableTDText)
								}
							}
						}
						used = false
						if !inline {
							XMLTable.WriteString(XMLIMGtail)
						}
					}
					tds++
				}
			}
			if inline && !used {
				XMLTable.WriteString(XMLIMGtail)
			}
			XMLTable.WriteString(XMLHeadTableTDEnd)
		}
		XMLTable.WriteString(XMLTableEndTR)
	}
	XMLTable.WriteString(XMLTableFooter)
	//serialization
	var rows []interface{}
	for _, row := range tableBody {
		for _, rowdata := range row {
			for _, rowEle := range rowdata.TData {
				if _, ok := rowEle.([][][]interface{}); !ok {
					if text, ok := rowEle.(*Text); ok {
						tColor := text.Color
						tSize := text.Size
						tWord := text.Words
						rows = append(rows, tColor, tSize, tSize, tWord)
					}
				}
			}
		}
	}

	//data fill in
	tabledata := fmt.Sprintf(XMLTable.String(), rows...)
	// fmt.Printf("table XML内容：%+v", tabledata)
	d.ReplaceXML(table.Tbname, tabledata, -1)
	return nil
}

func replaceTableToBuffer(table *Table) (string, error) {
	tableHead := table.TableHead
	tableBody := table.TableBody
	inline := table.Inline
	XMLTable := bytes.Buffer{}
	var Bused bool
	Bused = false
	//handle TableHead :Split with TableBody
	if tableHead != nil {
		//表格中的表格为无边框形式
		XMLTable.WriteString(XMLTableInTableHead)
		XMLTable.WriteString(XMLTableHeadTR)
		for thindex, rowdata := range tableHead {
			thw := fmt.Sprintf(XMLHeadTableTDBegin, strconv.FormatInt(int64(rowdata[thindex].TDW), 10))
			XMLTable.WriteString(thw)
			if inline {
				if table.Thcenter {
					XMLTable.WriteString(XMLHeadTableTDBegin2C)
				} else {
					XMLTable.WriteString(XMLHeadTableTDBegin2)
				}
			}
			for _, rowEle := range rowdata {
				if !inline {
					if table.Thcenter {
						XMLTable.WriteString(XMLHeadTableTDBegin2C)
					} else {
						XMLTable.WriteString(XMLHeadTableTDBegin2)
					}
				}
				if text, ok := rowEle.TData.(*Text); ok {
					color := text.Color
					size := text.Size
					word := text.Words
					var data string
					if text.IsCenter {
						if text.Isbold {
							data = fmt.Sprintf(XMLHeadtableTDTextBC, color, size, size, word)
						} else {
							data = fmt.Sprintf(XMLHeadtableTDTextC, color, size, size, word)
						}
					} else {
						if text.Isbold {
							data = fmt.Sprintf(XMLHeadtableTDTextB, color, size, size, word)
						} else {
							data = fmt.Sprintf(XMLHeadtableTDText, color, size, size, word)
						}
					}
					XMLTable.WriteString(data)
				}
				if !inline {
					XMLTable.WriteString(XMLIMGtail)
				}
			}
			if inline {
				XMLTable.WriteString(XMLIMGtail)
			}
			XMLTable.WriteString(XMLHeadTableTDEnd)
		}
		XMLTable.WriteString(XMLTableEndTR)
	} else {
		XMLTable.WriteString(XMLTableInTableNoHead)
	}

	//Generate formation
	for _, v := range tableBody {
		XMLTable.WriteString(XMLTableTR)

		for _, vv := range v {
			var ttd string
			if vv.TDM == 1 {
				ttd = fmt.Sprintf(XMLTableInTableMergeSTD, strconv.FormatInt(int64(vv.TDW), 10))
			} else if vv.TDM == 2 {
				ttd = fmt.Sprintf(XMLTableInTableMergeCTD, strconv.FormatInt(int64(vv.TDW), 10))
			} else {
				ttd = fmt.Sprintf(XMLTableInTableTD, strconv.FormatInt(int64(vv.TDW), 10))
			}

			XMLTable.WriteString(ttd)
			if vv.TDM < 2 {
				tds := 0
				if inline {
					XMLTable.WriteString(XMLTableTD2)

				}
				for _, vvv := range vv.TData {
					table, ok := vvv.(*Table)
					if !inline && !ok {
						XMLTable.WriteString(XMLTableTD2)
					}
					if ok {
						Bused = true
						tablestr, err := replaceTableToBuffer(table)
						if err != nil {
							return "", err
						}
						XMLTable.WriteString(tablestr)
						XMLTable.WriteString(XMLMagicFooter)
					} else {
						if text, ko := vvv.(*Text); ko {
							if text.IsCenter {
								if text.Isbold {
									XMLTable.WriteString(XMLHeadtableTDTextBC)
								} else {
									XMLTable.WriteString(XMLHeadtableTDTextC)
								}
							} else {
								if text.Isbold {
									XMLTable.WriteString(XMLHeadtableTDTextB)
								} else {
									fmt.Printf("model:%+v", vv)
									XMLTable.WriteString(XMLHeadtableTDText)
								}
							}
						}
						//not end with table
						Bused = false
						var next bool
						if tds < len(vv.TData)-1 {
							_, next = vv.TData[tds+1].(*Table)
						}

						if !inline {
							XMLTable.WriteString(XMLIMGtail)
						} else if inline && next {
							XMLTable.WriteString(XMLIMGtail)
						}
					}
					tds++
				}
				if inline && !Bused {
					XMLTable.WriteString(XMLIMGtail)
				}
			}
			XMLTable.WriteString(XMLHeadTableTDEnd)
		}
		XMLTable.WriteString(XMLTableEndTR)
	}
	XMLTable.WriteString(XMLTableFooter)
	//serialization
	var rows []interface{}

	for _, row := range tableBody {
		for _, rowdata := range row {
			for _, rowEle := range rowdata.TData {
				if _, ok := rowEle.([][][]interface{}); !ok {
					if text, ok := rowEle.(*Text); ok {
						tColor := text.Color
						tSize := text.Size
						tWord := text.Words
						rows = append(rows, tColor, tSize, tSize, tWord)
					}
				}
			}
		}
	}

	//data fill in
	tabledata := fmt.Sprintf(XMLTable.String(), rows...)

	return tabledata, nil
}

//NewTable create a table
func NewTable(d *Docx, tbname string, inline bool, tableBody [][]*TableTD, tableHead [][]*TableTHead, headCenter bool) (*Table, error) {
	table := &Table{}
	table.Tbname = tbname
	table.Inline = inline
	table.TableBody = tableBody
	table.TableHead = tableHead
	table.Thcenter = headCenter
	err := d.ReplaceTable(table)
	return table, err
}

func NewTableTD(tdata []interface{}, tdproperty map[string]interface{}) *TableTD {
	Tabletd := &TableTD{
		TData: tdata,
		TDBG:  0,
		TDW:   0,
		TDM:   0,
	}
	if tdproperty != nil {
		if tdproperty["tdbg"] != nil {
			Tabletd.TDBG = tdproperty["tdbg"].(int)
		}
		if tdproperty["tdw"] != nil {
			Tabletd.TDW = tdproperty["tdw"].(int)
		}
		if tdproperty["tdm"] != nil {
			Tabletd.TDM = tdproperty["tdm"].(int)
		}
	}
	return Tabletd
}

func NewText(words string) *Text {
	words, _ = encode(words)
	text := &Text{}
	text.Words = words
	text.Color = "000000"
	text.Size = "19"
	text.Isbold = false
	text.IsCenter = false
	return text
}

func (d *Docx) Write(ioWriter io.Writer) (err error) {
	w := zip.NewWriter(ioWriter)
	for _, file := range d.Files {
		var writer io.Writer
		var readCloser io.ReadCloser

		writer, err = w.Create(file.Name)
		if err != nil {
			return err
		}
		readCloser, err = file.Open()
		if err != nil {
			return err
		}
		if file.Name == "word/document.xml" {
			writer.Write([]byte(d.Content))
		} else {
			writer.Write(streamToByte(readCloser))
		}
	}
	w.Close()
	return
}

func (d *Docx) WriteToFile(path string) (err error) {
	var target *os.File
	target, err = os.Create(path)
	if err != nil {
		return
	}
	defer target.Close()
	err = d.Write(target)
	return
}

func streamToByte(stream io.Reader) []byte {
	buf := new(bytes.Buffer)
	buf.ReadFrom(stream)
	return buf.Bytes()
}

func encode(s string) (string, error) {
	var b bytes.Buffer
	enc := xml.NewEncoder(bufio.NewWriter(&b))
	if err := enc.Encode(s); err != nil {
		return s, err
	}
	output := strings.Replace(b.String(), "<string>", "", 1) // remove string tag
	output = strings.Replace(output, "</string>", "", 1)
	output = strings.Replace(output, "&#xD;&#xA;", "<w:br/>", -1) // \r\n => newline
	return output, nil
}
