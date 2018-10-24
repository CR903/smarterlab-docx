package docx

const (

	//XMLText == 正文
	XMLText = `<w:p>
		<w:r>
			<w:rPr>
				<w:color w:val="%s"/>
				<w:sz w:val="%s"/>
				<w:sz-cs w:val="%s"/>
			</w:rPr>
		<w:t>%s</w:t>
		</w:r>
  	</w:p>
`
	//XMLCenterText == 居中正文
	XMLCenterText = `<w:p>
		<w:r>
		<w:pPr>
			<w:jc w:val="center"/>
		</w:pPr>
		<w:rPr>
			<w:color w:val="%s"/>
			<w:sz w:val="%s"/>
			<w:sz-cs w:val="%s"/>
		</w:rPr>
			<w:t>%s</w:t>
		</w:r>
	</w:p>
`
	//XMLCenterBoldText 居中粗体
	XMLCenterBoldText = `<w:p>
		<w:r>
			<w:pPr>
				<w:jc w:val="center"/>
			</w:pPr>
			<w:rPr>
				<w:b/>
				<w:b-cs/>
				<w:color w:val="%s"/>
				<w:sz w:val="%s"/>
				<w:sz-cs w:val="%s"/>
			</w:rPr>
			<w:t>%s</w:t>
		</w:r>
	</w:p>
`
	//XMLBoldText ==粗体
	XMLBoldText = `<w:p>
		<w:r>
			<w:rPr>
				<w:b/>
				<w:b-cs/>
				<w:color w:val="%s"/>
				<w:sz w:val="%s"/>
			</w:rPr>
			<w:t>%s</w:t>
		</w:r>
	</w:p>
`
	//XMLInlineText == 不换行的正文
	XMLInlineText = `<w:r>
		<w:t>%s</w:t>
	</w:r>
`
	//XMLFontStyle defines fontStyle
	XMLFontStyle = `<w:pPr>
		<w:pStyle w:val="%s"/>
		<w:jc w:val="center"/>
		<w:textAlignment w:val="center"/>
	</w:pPr>
`
	//XMLTableHead ...
	XMLTableHead = `<w:tbl>
	<w:tblPr>
		<w:tblStyle w:val="ableGrid"/>
		<w:tblW w:w="0" w:type="pct"/>
	</w:tblPr>
`
	//XMLTableNoHead == 没有表头的样式把table top line remove掉
	XMLTableNoHead = `<w:tbl>
	<w:tblPr>
		<w:tblStyle w:val="a3"/>
		<w:tblW w:w="0" w:type="auto"/>			
		<w:tblLook w:firstColumn="1" w:firstRow="1" w:lastColumn="0" w:lastRow="0" w:noHBand="0" w:noVBand="1" w:val="04A0"/>
	</w:tblPr>
`
	//XMLTableInTableHead == 表中表的样式头
	XMLTableInTableHead = `<w:tbl>
	<w:tblPr>
		<w:tblStyle w:val="a3"/>
		<w:tblW w:w="0" w:type="auto"/>				
		<w:tblLook w:firstColumn="1" w:firstRow="1" w:lastColumn="0" w:lastRow="0" w:noHBand="0" w:noVBand="1" w:val="04A0"/>
	</w:tblPr>
`
	//XMLTableInTableNoHead ...
	XMLTableInTableNoHead = `<w:tbl>
		<w:tblPr>
		<w:tblStyle w:val="a3"/>
		<w:tblW w:w="0" w:type="auto"/>				
		<w:tblLook w:firstColumn="1" w:firstRow="1" w:lastColumn="0" w:lastRow="0" w:noHBand="0" w:noVBand="1" w:val="04A0"/>
	</w:tblPr>
`
	//XMLTableTR ...
	XMLTableTR = `<w:tr w14:paraId="7F912C80" w14:textId="77777777" w:rsidR="00102180" w:rsidTr="00102180">
`
	//XMLTableHeadTR ...
	XMLTableHeadTR = `<w:tr w14:paraId="7F912C80" w14:textId="77777777" w:rsidR="00102180" w:rsidTr="00102180">
	<w:trPr>
		<w:trHeight w:val="auto"/>
	</w:trPr>
`
	//XMLTableTD ...
	XMLTableTD = `<w:tc>
	<w:tcPr>
		<w:tcW w:w="%s" w:type="dxa"/>
		<w:gridSpan w:val="%s"/>
	</w:tcPr>
`
	// XMLTableMergeSTD ...
	XMLTableMergeSTD = `<w:tc>
	<w:tcPr>
		<w:tcW w:w="%s" w:type="dxa"/>
		<w:gridSpan w:val="%s"/>
		<w:vMerge w:val="restart"/>
	</w:tcPr>
`
	// XMLTableMergeCTD ...
	XMLTableMergeCTD = `<w:tc>
	<w:tcPr>
		<w:tcW w:w="%s" w:type="dxa"/>
		<w:gridSpan w:val="%s"/>
		<w:vMerge/>
	</w:tcPr>
	<w:p w14:paraId="%s" w14:textId="%s" w:rsidR="%s" w:rsidRDefault="%s"/>
`

	//XMLTableInTableTD ...
	XMLTableInTableTD = `<w:tc>
	<w:tcPr>
		<w:tcW w:w="%s" w:type="dxa"/>
	</w:tcPr>
`
	// XMLTableInTableMergeSTD ...
	XMLTableInTableMergeSTD = `<w:tc>
	<w:tcPr>
		<w:tcW w:w="%s" w:type="dxa"/>
		<w:vMerge w:val="restart"/>
	</w:tcPr>
`
	// XMLTableInTableMergeCTD ...
	XMLTableInTableMergeCTD = `<w:tc>
	<w:tcPr>
		<w:tcW w:w="%s" w:type="dxa"/>
		<w:vMerge/>
	</w:tcPr>
	<w:p w14:paraId="%s" w14:textId="%s" w:rsidR="%s" w:rsidRDefault="%s"/>
`

	//XMLTableTD2 ...
	XMLTableTD2 = `<w:p>
	<w:pPr>
		<w:tabs>
			<w:tab w:val="center" w:pos="1312"/>
		</w:tabs>
		<w:textAlignment w:val="center"/>
	</w:pPr>
`
	//XMLHeadTableTDBegin ...
	XMLHeadTableTDBegin = `<w:tc>
	<w:tcPr>
		<w:tcW w:w="%s" w:type="dxa"/>
	</w:tcPr>
`
	//XMLHeadTableInTableTDBegin ...
	XMLHeadTableInTableTDBegin = `<w:tc>
	<w:tcPr>
		<w:tcW w:w="%s" w:type="dxa"/>
	</w:tcPr>
`
	//XMLHeadTableTDBegin2 ...
	XMLHeadTableTDBegin2 = `<w:p>
	<w:pPr>
		<w:textAlignment w:val="left"/>
	</w:pPr>
`

	//XMLHeadTableTDBegin2C ...
	XMLHeadTableTDBegin2C = `<w:p>
	<w:pPr>
		<w:textAlignment w:val="center"/>
	</w:pPr>
`

	//XMLHeadtableTDTextB ...
	XMLHeadtableTDTextB = `<w:r>
		<w:rPr>
			<w:b/>
			<w:b-cs/>
			<w:color w:val="%s"/>
			<w:sz w:val="%s"/>
			<w:sz-cs w:val="%s"/>
		</w:rPr>
		<w:t>%s</w:t>
	</w:r>
`
	//XMLHeadtableTDTextC ...
	XMLHeadtableTDTextC = `<w:pPr>
		<w:textAlignment w:val="center"/>
	</w:pPr>
	<w:r>
		<w:rPr>
			<w:color w:val="%s"/>
			<w:sz w:val="%s"/>
			<w:sz-cs w:val="%s"/>
		</w:rPr>
		<w:t>%s</w:t>
	</w:r>
 `
	//XMLHeadtableTDTextBC ...
	XMLHeadtableTDTextBC = `<w:pPr>
		<w:textAlignment w:val="center"/>
	</w:pPr>
	<w:r>
		<w:rPr>
			<w:b/>
			<w:color w:val="%s"/>
			<w:sz w:val="%s"/>
			<w:sz-cs w:val="%s"/>
		</w:rPr>
		<w:t>%s</w:t>
	</w:r>
 `
	//XMLHeadtableTDText ...
	XMLHeadtableTDText = `<w:pPr>
		<w:textAlignment w:val="left"/>
	</w:pPr>
	<w:r>
	<w:rPr>
		<w:rFonts w:hint="fareast"/>
		<w:color w:val="%s"/>
		<w:sz w:val="%s"/>
		<w:sz-cs w:val="%s"/>
	</w:rPr>
		<w:t>%s</w:t>
	</w:r>
`

	// XMLTableGridBegin ...
	XMLTableGridBegin = `<w:tblGrid>
	`

	// XMLTableGridCol ...
	XMLTableGridCol = `<w:gridCol w:w="%s" />
	`

	// XMLTableGridEnd ...
	XMLTableGridEnd = `
	</w:tblGrid>`

	//XMLHeadTableTDEnd ...
	XMLHeadTableTDEnd = `
	</w:tc>
`
	//XMLTableEndTR ...
	XMLTableEndTR = `</w:tr>
`
	//XMLMagicFooter  HACK:I struggle for a long time,at last ,I find it is necessary,and don't konw why.
	XMLMagicFooter = `<w:p>
		<w:pPr>
			<w:tabs>
				<w:tab w:val="center" w:pos="1312"/>
			</w:tabs>
			<w:textAlignment w:val="auto"/>
		</w:pPr>
		<w:r>
			<w:t></w:t>
		</w:r>
	</w:p>
`

	//XMLTableFooter ...
	XMLTableFooter = `
	</w:tbl>
`

	//XMLIMGtail ...
	XMLIMGtail = `</w:p>
`
	//XMLBr == 换行
	XMLBr = `<w:p/>
`
)
