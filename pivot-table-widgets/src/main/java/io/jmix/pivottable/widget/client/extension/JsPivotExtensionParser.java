/*
 * Copyright 2021 Haulmont.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package io.jmix.pivottable.widget.client.extension;

import com.google.gwt.core.client.JavaScriptObject;
import com.google.gwt.dom.client.Element;
import com.google.gwt.i18n.client.DateTimeFormat;

public class JsPivotExtensionParser extends JavaScriptObject {

    protected JsPivotExtensionParser() {
    }

    public static native JsPivotExtensionParser create() /*-{
        return {
           boldClassNames: ['pvtTotal rowTotal', 'pvtTotal colTotal', 'pvtGrandTotal'],
           getTable: function(pivotElement) {
                var tableElements = pivotElement.getElementsByClassName('pvtTable');
                if (tableElements.length === 0) {
                    return;
                }
                return tableElements[0];
           },
           getColsRowsNumber: function(table, attributeName) {
                return table.attributes[attributeName] ? table.attributes[attributeName].value : null;
           },
           getRows: function(tableRows, appendIndex) {
                var modelRows = [];
                for (var i = 0; i < tableRows.length; i++) {
                    var modelRow = this.convertRowToModel(tableRows[i], i + appendIndex);
                    modelRows.push(modelRow);
                }
                return modelRows;
           },
           convertRowToModel: function(tableRow, rowNumber) {
                var modelRow = {};
                modelRow.tableRowNumber = rowNumber;
                modelRow.cells = [];

                for (var j = 0; j < tableRow.cells.length; j++) {
                    var cell = tableRow.cells[j];
                    var modelCell = {};

                    // check for bold
                    if (cell.nodeName === 'TH' || this.boldClassNames.indexOf(cell.className) > -1) {
                        modelCell.isBold = true;
                    } else {
                        modelCell.isBold = false;
                    }

                    if (cell.className) {
                        modelCell.className = cell.className;
                    }

                    modelCell.colSpan = cell.colSpan;
                    modelCell.rowSpan = cell.rowSpan;

                    if (cell.outerText) {
                        modelCell.value = cell.outerText;
                    } else if (cell.innerText) {
                        modelCell.value = cell.innerText;
                    }

                    if (cell.attributes['data-value']) {
                        if (cell.attributes['data-value'].value !== 'null'
                                && cell.attributes['data-value'].value !== "") {
                            modelCell.value = cell.attributes['data-value'].value;
                        }
                    }

                    modelCell.type = this.getCellType(modelCell.value);

                    modelRow.cells.push(modelCell);
                }
                return modelRow;
           },
           getCellType: function(value) {
                var isDate = $entry(function(value, format){
                    return @io.jmix.pivottable.widget.client.extension.JsPivotExtensionParser::isDate(Ljava/lang/String;Ljava/lang/String;)(value, format);
                });

                if (!isNaN(value)) {
                    return value % 1 == 0 ? 'INTEGER' : 'DOUBLE';
                } else if (isDate(value, this.dateTimeFormat)) {
                    return 'DATE_TIME';
                } else if (isDate(value, this.dateFormat)) {
                    return 'DATE';
                } else if (isDate(value, this.timeFormat)) {
                    return 'TIME';
                } else {
                   return 'STRING';
                }
           }
        };
    }-*/;

    public final native void setDateTimeParseFormat(JsPivotExtensionParser parser, String format) /*-{
        parser.dateTimeParseFormat = format;
    }-*/ ;

    public final native void setDateParseFormat(JsPivotExtensionParser parser, String format) /*-{
        parser.dateParseFormat = format;
    }-*/ ;

    public final native void setTimeParseFormat(JsPivotExtensionParser parser, String format) /*-{
        parser.timeParseFormat = format;
    }-*/ ;

    public final native String parsePivotTableToJson(JsPivotExtensionParser parser, Element pivotElement) /*-{
        var table = parser.getTable(pivotElement);

        var resultObject = {};
        resultObject.dataNumRows = parser.getColsRowsNumber(table, 'data-numrows');
        resultObject.dataNumCols = parser.getColsRowsNumber(table, 'data-numcols');

        var tableHead = table.tHead;
        resultObject.headRows = parser.getRows(tableHead.rows, 0);

        var tableBody = table.tBodies[0];
        resultObject.bodyRows = parser.getRows(tableBody.rows, resultObject.headRows.length);

        return JSON.stringify(resultObject);
    }-*/;

    protected static boolean isDate(String value, String format) {
        if (value == null || value.isEmpty()) {
            return false;
        }

        if (format == null || format.isEmpty()) {
            return false;
        }

        DateTimeFormat dateTimeFormat = DateTimeFormat.getFormat(format);
        try {
            dateTimeFormat.parse(value);
            return true;
        } catch (Exception e) {
            return false;
        }
    }
}
