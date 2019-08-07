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
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.moqui.BaseArtifactException
import org.moqui.impl.context.ExecutionContextImpl
import org.moqui.impl.screen.ScreenDefinition
import org.moqui.impl.screen.ScreenForm
import org.moqui.impl.screen.ScreenRenderImpl
import org.moqui.impl.screen.ScreenWidgetRender
import org.moqui.impl.screen.ScreenWidgets
import org.moqui.util.ContextStack
import org.moqui.util.MNode
import org.slf4j.Logger
import org.slf4j.LoggerFactory

@CompileStatic
class ScreenWidgetRenderExcel implements ScreenWidgetRender {
    private static final Logger logger = LoggerFactory.getLogger(ScreenWidgetRenderExcel.class)
    private static final String workbookFieldName = "WidgetRenderXSSFWorkbook"

    ScreenWidgetRenderExcel() { }

    @Override
    void render(ScreenWidgets widgets, ScreenRenderImpl sri) {
        ContextStack cs = sri.ec.contextStack
        cs.push()
        try {
            cs.sri = sri
            OutputStream os = sri.getOutputStream()

            XSSFWorkbook wb = (XSSFWorkbook) cs.getByString(workbookFieldName)
            boolean createdWorkbook = false
            if (wb == null) {
                wb = new XSSFWorkbook()
                createdWorkbook = true
                cs.put(workbookFieldName, wb)
            }

            MNode widgetsNode = widgets.widgetsNode
            if (widgetsNode.name == "screen") widgetsNode = widgetsNode.first("widgets")
            renderSubNodes(widgetsNode, sri, wb)

            if (createdWorkbook) {
                if (wb.getNumberOfSheets() > 0) {
                    // write file to stream
                    wb.write(os)
                    wb.close()
                    os.close()
                } else {
                    throw new BaseArtifactException("No sheets rendered")
                }
            }
        } finally {
            cs.pop()
        }
    }

    static void renderSubNodes(MNode widgetsNode, ScreenRenderImpl sri, XSSFWorkbook workbook) {
        ExecutionContextImpl eci = sri.ec
        ScreenDefinition sd = sri.getActiveScreenDef()

        // iterate over child elements to find and render form-list
        // recursive renderSubNodes() call for: container-box (box-body, box-body-nopad), container-row (row-col)
        ArrayList<MNode> childList = widgetsNode.getChildren()
        int childListSize = childList.size()
        for (int i = 0; i < childListSize; i++) {
            MNode childNode = (MNode) childList.get(i)
            String nodeName = childNode.getName()
            if ("form-list".equals(nodeName)) {
                ScreenForm form = sd.getForm(childNode.attribute("name"))
                MNode formNode = form.getOrCreateFormNode()
                String formName = eci.resourceFacade.expandNoL10n(formNode.attribute("name"), null)

                ScreenForm.FormInstance formInstance = form.getFormInstance()

                XSSFSheet sheet = workbook.createSheet(formName)
                FormListExcelRender fler = new FormListExcelRender(formInstance, eci)
                fler.renderSheet(sheet)
            } else if ("section".equals(nodeName)) {
                // nest into section by calling renderSection() so conditions, actions are run (skip section-iterate)
                sri.renderSection(childNode.attribute("name"))
            } else if ("container-box".equals(nodeName)) {
                MNode boxBody = childNode.first("box-body")
                if (boxBody != null) renderSubNodes(boxBody, sri, workbook)
                MNode boxBodyNopad = childNode.first("box-body-nopad")
                if (boxBodyNopad != null) renderSubNodes(boxBodyNopad, sri, workbook)
            } else if ("container-row".equals(nodeName)) {
                MNode rowCol = childNode.first("row-col")
                if (rowCol != null) renderSubNodes(rowCol, sri, workbook)
            }
            // NOTE: other elements ignored, including section-iterate (outside intended use case for Excel render
        }
    }
}
