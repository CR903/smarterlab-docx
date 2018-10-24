package main

import (
	"fmt"

	"github.com/CR903/smarterlab-docx/docx"
)

func main() {

	r, err := docx.ReadDocxFile("template.docx")
	if err != nil {
		panic(err)
	}

	docxObj := r.Editable()
	docxObj.Replace("DETECTNUMBER", "JC2018001", -1)

	row1TD1Key := docx.NewTableTD([]interface{}{docx.NewText("委托单位")}, map[string]interface{}{"tdbg": 1, "tdw": 2095})
	row1TD1Value := docx.NewTableTD([]interface{}{docx.NewText("立为科技")}, map[string]interface{}{"tdbg": 1, "tdw": 2095})
	row1TD2Key := docx.NewTableTD([]interface{}{docx.NewText("联系人")}, map[string]interface{}{"tdbg": 1, "tdw": 2095})
	row1TD2Value := docx.NewTableTD([]interface{}{docx.NewText("联系人")}, map[string]interface{}{"tdbg": 1, "tdw": 2095})

	row2TD1Key := docx.NewTableTD([]interface{}{docx.NewText("委托单位地址")}, nil)
	row2TD1Value := docx.NewTableTD([]interface{}{docx.NewText("地址1")}, nil)
	row2TD2Key := docx.NewTableTD([]interface{}{docx.NewText("电 话")}, nil)
	row2TD2Value := docx.NewTableTD([]interface{}{docx.NewText("13600000000")}, nil)

	row3TD1Key := docx.NewTableTD([]interface{}{docx.NewText("付款单位")}, nil)
	row3TD1Value := docx.NewTableTD([]interface{}{docx.NewText("")}, nil)
	row3TD2Key := docx.NewTableTD([]interface{}{docx.NewText("税号")}, nil)
	row3TD2Value := docx.NewTableTD([]interface{}{docx.NewText("")}, nil)

	row4TD1Key := docx.NewTableTD([]interface{}{docx.NewText("受测单位")}, nil)
	row4TD1Value := docx.NewTableTD([]interface{}{docx.NewText("垃圾场")}, nil)
	row4TD2Key := docx.NewTableTD([]interface{}{docx.NewText("联系人")}, nil)
	row4TD2Value := docx.NewTableTD([]interface{}{docx.NewText("赵六")}, nil)

	row5TD1Key := docx.NewTableTD([]interface{}{docx.NewText("通讯地址")}, nil)
	row5TD1Value := docx.NewTableTD([]interface{}{docx.NewText("地址4")}, nil)
	row5TD2Key := docx.NewTableTD([]interface{}{docx.NewText("电 话")}, nil)
	row5TD2Value := docx.NewTableTD([]interface{}{docx.NewText("13996966666")}, nil)

	row6TD1Key := docx.NewTableTD([]interface{}{docx.NewText("样品名称")}, nil)
	row6TD1Value := docx.NewTableTD([]interface{}{docx.NewText("水")}, nil)
	row6TD2Key := docx.NewTableTD([]interface{}{docx.NewText("样品数量")}, nil)
	row6TD2Value := docx.NewTableTD([]interface{}{docx.NewText("12.00")}, nil)

	row7TD1Key := docx.NewTableTD([]interface{}{docx.NewText("样品名称")}, nil)
	row7TD1Value := docx.NewTableTD([]interface{}{docx.NewText("气")}, nil)
	row7TD2Key := docx.NewTableTD([]interface{}{docx.NewText("样品数量")}, nil)
	row7TD2Value := docx.NewTableTD([]interface{}{docx.NewText("12.00")}, nil)

	row8TD1Key := docx.NewTableTD([]interface{}{docx.NewText("检测项目")}, map[string]interface{}{"tdm": 1})
	row8TD1Value := docx.NewTableTD([]interface{}{docx.NewText("水：砷, ph")}, map[string]interface{}{"tdbg": 3})

	row81TD1Key := docx.NewTableTD([]interface{}{docx.NewText("")}, map[string]interface{}{"tdm": 2})
	row81TD1Value := docx.NewTableTD([]interface{}{docx.NewText("气:测试检测项, 四氯化碳")}, map[string]interface{}{"tdbg": 3})

	row9TD1Key := docx.NewTableTD([]interface{}{docx.NewText("检测标准/检测方法")}, nil)
	row9TD1Value := docx.NewTableTD([]interface{}{docx.NewText("标准1/标准1, 标准1/标准1, 测试检测标准一号/测试检测标准一号, 标准1/标准1")}, map[string]interface{}{"tdbg": 3, "tdw": 0})

	row10TD1Key := docx.NewTableTD([]interface{}{docx.NewText("具体检测项目及执行标准/检测方法以检测项目附件单为准（需双方签字确认）")}, map[string]interface{}{"tdbg": 4, "tdw": 0})

	row11TD1Key := docx.NewTableTD([]interface{}{docx.NewText("服务类型")}, nil)
	row11TD1Value := docx.NewTableTD([]interface{}{docx.NewText("☑ 标准服务：7个工作日"), docx.NewText("☐ 加急服务：3.5个工作日 加收100%附加费")}, map[string]interface{}{"tdbg": 3, "tdw": 0})

	row12TD1Key := docx.NewTableTD([]interface{}{docx.NewText("报告份数")}, nil)
	row12TD1Value := docx.NewTableTD([]interface{}{docx.NewText("")}, nil)
	row12TD2Key := docx.NewTableTD([]interface{}{docx.NewText("预计完成时间")}, nil)
	row12TD2Value := docx.NewTableTD([]interface{}{docx.NewText("2018-10-31")}, nil)

	row13TD1Key := docx.NewTableTD([]interface{}{docx.NewText("取报告方式")}, nil)
	row13TD1Value := docx.NewTableTD([]interface{}{docx.NewText("☑ 自取"), docx.NewText("☐ 快递")}, nil)
	row13TD2Key := docx.NewTableTD([]interface{}{docx.NewText("总费用")}, nil)
	row13TD2Value := docx.NewTableTD([]interface{}{docx.NewText("200.000000")}, nil)

	row14TD1Key := docx.NewTableTD([]interface{}{docx.NewText("说明："),
		docx.NewText("1）委托单位如对检测方法有特殊要求，请在执行标准/检测方法中详细说明。"),
		docx.NewText("2）委托单位如对盖资质章有要求，请在备注中说明。"),
		docx.NewText("3）若双方另有其他要求可附页说明。")}, map[string]interface{}{"tdbg": 4, "tdw": 0})

	row15TD1Key := docx.NewTableTD([]interface{}{docx.NewText("备注："),
		docx.NewText("")}, map[string]interface{}{"tdbg": 4, "tdw": 0})

	table := [][]*docx.TableTD{
		{row1TD1Key, row1TD1Value, row1TD2Key, row1TD2Value},
		{row2TD1Key, row2TD1Value, row2TD2Key, row2TD2Value},
		{row3TD1Key, row3TD1Value, row3TD2Key, row3TD2Value},
		{row4TD1Key, row4TD1Value, row4TD2Key, row4TD2Value},
		{row5TD1Key, row5TD1Value, row5TD2Key, row5TD2Value},
		{row6TD1Key, row6TD1Value, row6TD2Key, row6TD2Value},
		{row7TD1Key, row7TD1Value, row7TD2Key, row7TD2Value},
		{row8TD1Key, row8TD1Value},
		{row81TD1Key, row81TD1Value},
		{row9TD1Key, row9TD1Value},
		{row10TD1Key},
		{row11TD1Key, row11TD1Value},
		{row12TD1Key, row12TD1Value, row12TD2Key, row12TD2Value},
		{row13TD1Key, row13TD1Value, row13TD2Key, row13TD2Value},
		{row14TD1Key},
		{row15TD1Key},
	}
	_, err1 := docx.NewTable(docxObj, "DETECTCONTENT", false, table, nil, false)
	if err1 != nil {
		panic(err1)
	}
	docxObj.WriteToFile("template_new.docx")
	r.Close()
	fmt.Println("Success")
}
