/*
 * Copyright 2009 the original author or authors.
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

package org.jggug.kobo.gexcelapi

import groovy.xml.MarkupBuilder

class CellRangeToHtmlConverter {

    CellRange range

    CellRangeToHtmlConverter(CellRange range) {
        this.range = range
    }

    String toHtml(String headTitle, String charset) {
        def writer = new StringWriter()
        new MarkupBuilder(writer).html {
            head {
                meta('http-equiv': 'Content-Type', content: "text/html; charset=${charset}")
                title(headTitle)
                style(type:'text/css') {
                    mkp.yieldUnescaped('''\
                        |<!--
                        |table {
                        |   border-collapse: collapse;
                        |   border: 2px solid black;
                        |}
                        |tr:hover {
                        |   background: #ffb;
                        |}
                        |td, th {
                        |   border: 1px solid black;
                        |}
                        |td:hover {
                        |   background: #ee9;
                        |}
                        |th {
                        |   background: #ddd;
                        |}
                        |//-->
                        | '''.stripMargin('|'))
                }
            }
            body {
                table {
                    tr {
                        th ""
                        range.first().each { cell -> th cell?.columnLabel }
                    }
                    range.each { row ->
                        tr {
                            th row.first()?.rowLabel
                            row.each { cell ->
                                def region = range.sheet.getEnclosingMergedRegion(cell)
                                if (region) {
                                    if (!region.isFirstCell(cell)) return
                                    def attrs = [colspan:region.width, rowspan:region.height]
                                    td attrs, cell?.value
                                } else {
                                    td cell?.value
                                }
                            }
                        }
                    }
                }
            }
        }
        return writer.toString()
    }

}
