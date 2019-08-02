/*
 * This software is in the public domain under CC0 1.0 Universal plus a
 * Grant of Patent License.
 *
 * To the extent possible under law, the author(s) have dedicated all
 * copyright and related and neighboring rights to this software to the
 * public domain worldwide. This software is distributed without any
 * warranty.
 *
 * You should have received a copy of the CC0 Public Domain Dedication
 * along with this software (see the LICENSE.md file). If not, see
 * <http://creativecommons.org/publicdomain/zero/1.0/>.
 */
package org.moqui.poi

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.DateFormatConverter
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.moqui.impl.context.ExecutionContextImpl
import org.moqui.impl.screen.ScreenForm
import org.moqui.util.MNode
import org.moqui.util.StringUtilities

import java.text.DateFormat

/**
 * Generate an Excel (XLSX) file with form-list based output similar to CSV export.
 *
 * Note that this relies on internal classes in Moqui Framework, ie an unstable API.
 */
@CompileStatic
class FormListExcelRender {
    // ignore the same field elements as the DefaultScreenMacros.csv.ftl file
    protected static Set<String> ignoreFieldElements = new HashSet<>(["ignored", "hidden", "submit", "image", "label",
            "date-find", "file", "password", "range-find", "reset"])

    protected ScreenForm.FormInstance formInstance
    protected ExecutionContextImpl eci

    FormListExcelRender(ScreenForm.FormInstance formInstance, ExecutionContextImpl eci) {
        this.formInstance = formInstance
        this.eci = eci
    }

    void render(OutputStream os) {
        // get form-list data
        ScreenForm.FormListRenderInfo formListInfo = formInstance.makeFormListRenderInfo()
        MNode formNode = formListInfo.formNode
        String formName = eci.resourceFacade.expandNoL10n(formNode.attribute("name"), null)
        ArrayList<ArrayList<MNode>> formListColumnList = formListInfo.getAllColInfo()
        int numColumns = formListColumnList.size()

        // make POI workbook and sheet
        XSSFWorkbook wb = new XSSFWorkbook()
        XSSFSheet sheet = wb.createSheet(formName)

        CellStyle headerStyle = makeHeaderStyle(wb)
        CellStyle rowDefaultStyle = makeRowDefaultStyle(wb)
        CellStyle rowDateTimeStyle = makeRowDateTimeStyle(wb)
        CellStyle rowDateStyle = makeRowDateStyle(wb)
        CellStyle rowTimeStyle = makeRowTimeStyle(wb)
        CellStyle rowNumberStyle = makeRowNumberStyle(wb)
        CellStyle rowCurrencyStyle = makeRowCurrencyStyle(wb)

        int rowNum = 0

        // ========== header row
        XSSFRow headerRow = sheet.createRow(rowNum++)
        int colIndex = 0
        for (int i = 0; i < numColumns; i++) {
            ArrayList<MNode> columnFieldList = (ArrayList<MNode>) formListColumnList.get(i)
            for (int j = 0; j < columnFieldList.size(); j++) {
                MNode fieldNode = (MNode) columnFieldList.get(j)
                XSSFCell headerCell = headerRow.createCell(colIndex++)
                headerCell.setCellStyle(headerStyle)

                MNode headerField = fieldNode.first("header-field")
                MNode defaultField = fieldNode.first("default-field")
                String headerTitle = headerField != null ? headerField.attribute("title") : null
                if (headerTitle == null) headerTitle = defaultField != null ? defaultField.attribute("title") : null
                if (headerTitle == null) headerTitle = StringUtilities.camelCaseToPretty(fieldNode.attribute("name"))
                headerCell.setCellValue(headerTitle)
            }
        }

        // ========== line rows

        // should be ArrayList<Map<String, Object>> from call to AggregationUtil.aggregateList()
        Object listObject = formListInfo.getListObject(false)
        ArrayList<Map<String, Object>> listArray
        if (listObject instanceof ArrayList) {
            listArray = (ArrayList<Map<String, Object>>) listObject
        } else {
            throw new IllegalArgumentException("List object from FormListRenderInfo was not an ArrayList as expected, is ${listObject?.getClass()?.getName()}")
        }
        int listArraySize = listArray.size()

        for (int listIdx = 0; listIdx < listArraySize; listIdx++) {
            Map<String, Object> rowMap = (Map<String, Object>) listArray.get(listIdx)

            XSSFRow listRow = sheet.createRow(rowNum++)
            colIndex = 0

            for (int i = 0; i < numColumns; i++) {
                ArrayList<MNode> columnFieldList = (ArrayList<MNode>) formListColumnList.get(i)
                for (int j = 0; j < columnFieldList.size(); j++) {
                    MNode fieldNode = (MNode) columnFieldList.get(j)
                    String fieldName = fieldNode.attribute("name")

                    MNode defaultField = fieldNode.first("default-field")
                    ArrayList<MNode> childList = defaultField.getChildren()
                    int childListSize = childList.size()

                    XSSFCell listCell = listRow.createCell(colIndex++)

                    if (childListSize == 1) {
                        // use data specific cell type and style
                        MNode widgetNode = (MNode) childList.get(0)
                        String widgetType = widgetNode.getName()

                        // TODO style by widget type, etc
                        listCell.setCellStyle(rowDefaultStyle)

                        Object curValue = getFieldValue(fieldNode, widgetNode, rowMap)

                        // cell type options are _NONE, BLANK, BOOLEAN, ERROR, FORMULA, NUMERIC, STRING
                        // cell value options are: boolean, Date, Calendar, double, String, RichTextString
                        if (curValue instanceof String) {
                            listCell.setCellType(CellType.STRING)
                            listCell.setCellValue((String) curValue)
                        } else if (curValue instanceof Number) {
                            listCell.setCellType(CellType.NUMERIC)
                            listCell.setCellValue(((Number) curValue).doubleValue())
                        } else if (curValue instanceof Date) {
                            listCell.setCellType(CellType.STRING)
                            listCell.setCellValue((Date) curValue)
                        } else if (curValue != null) {
                            listCell.setCellType(CellType.STRING)
                            listCell.setCellValue(curValue.toString())
                        }
                    } else {
                        // always use string with values from all child elements
                        StringBuilder cellSb = new StringBuilder()

                        for (int childIdx = 0; childIdx < childListSize; childIdx++) {
                            MNode widgetNode = (MNode) childList.get(childIdx)
                            Object curValue = getFieldValue(fieldNode, widgetNode, rowMap)
                            if (curValue != null) {
                                if (curValue instanceof String) {
                                    cellSb.append((String) curValue)
                                } else {
                                    String format = widgetNode.attribute("format")
                                    String textFormat = widgetNode.attribute("text-format")
                                    if (textFormat != null && !textFormat.isEmpty()) format = textFormat
                                    cellSb.append(eci.l10nFacade.format(curValue, format))
                                }
                            }
                            if (childIdx < (childListSize - 1)) cellSb.append('\n')
                        }

                        listCell.setCellStyle(rowDefaultStyle)
                        listCell.setCellType(CellType.STRING)
                        listCell.setCellValue(cellSb.toString())
                    }
                }
            }
        }



        // TODO? ========== footer row


        // ========== output the result
        wb.write(os)
        os.close()
    }

    Object getFieldValue(MNode fieldNode, MNode widgetNode, Map<String, Object> rowMap) {
        String widgetType = widgetNode.getName()
        if (ignoreFieldElements.contains(widgetType)) return null
        String fieldName = fieldNode.attribute("name")

        // TODO this whole section for the various widget types...
        // similar logic to widget types in DefaultScreenMacros.csv.ftl
        Object value = null
        if ("display".equals(widgetType)) {

        } else if ("display-entity".equals(widgetType)) {

        } else if ("display".equals(widgetType)) {

        } else {
            value = rowMap.get(fieldName)
        }

        return value
    }

    CellStyle makeHeaderStyle(Workbook wb) {
        Font headerFont = wb.createFont()
        headerFont.setBold(true)

        CellStyle style = createBorderedStyle(wb)
        style.setAlignment(HorizontalAlignment.CENTER)
        style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex())
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND)
        style.setFont(headerFont)

        return style
    }

    CellStyle makeRowDefaultStyle(Workbook wb) {
        Font rowFont = wb.createFont()

        CellStyle style = createBorderedStyle(wb)
        style.setAlignment(HorizontalAlignment.LEFT)
        style.setFont(rowFont)

        return style
    }

    CellStyle makeRowDateTimeStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createBorderedStyle(wb)
        style.setAlignment(HorizontalAlignment.CENTER)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat(DateFormatConverter.convert(Locale.US, "yyyy-MM-dd HH:mm")))

        return style
    }
    CellStyle makeRowDateStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createBorderedStyle(wb)
        style.setAlignment(HorizontalAlignment.CENTER)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat(DateFormatConverter.convert(Locale.US, "yyyy-MM-dd")))

        return style
    }
    CellStyle makeRowTimeStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createBorderedStyle(wb)
        style.setAlignment(HorizontalAlignment.CENTER)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat(DateFormatConverter.convert(Locale.US, "HH:mm:ss")))

        return style
    }

    CellStyle makeRowNumberStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createBorderedStyle(wb)
        style.setAlignment(HorizontalAlignment.CENTER)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat("#,##0.######"))

        return style
    }
    CellStyle makeRowCurrencyStyle(Workbook wb) {
        DataFormat df = wb.createDataFormat()
        Font rowFont = wb.createFont()

        CellStyle style = createBorderedStyle(wb)
        style.setAlignment(HorizontalAlignment.RIGHT)
        style.setFont(rowFont)
        style.setDataFormat(df.getFormat('#,##0.00_);[Red](#,##0.00)'))

        return style
    }

    CellStyle createBorderedStyle(Workbook wb){
        BorderStyle thin = BorderStyle.THIN
        short black = IndexedColors.BLACK.getIndex()

        CellStyle style = wb.createCellStyle()
        style.setBorderRight(thin)
        style.setRightBorderColor(black)
        style.setBorderBottom(thin)
        style.setBottomBorderColor(black)
        style.setBorderLeft(thin)
        style.setLeftBorderColor(black)
        style.setBorderTop(thin)
        style.setTopBorderColor(black)
        return style
    }
}
