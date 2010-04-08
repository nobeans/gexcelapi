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

class CellLabelIteratorTest extends GroovyTestCase {

    void testIterator_label() {
        assert new CellLabelIterator("A1", "B3").collect { it } == ["A1", "A2", "A3", "B1", "B2", "B3"]
        assert new CellLabelIterator("A2", "B3").collect { it } == ["A2", "A3", "B2", "B3"]
        assert new CellLabelIterator("A2", "C3").collect { it } == ["A2", "A3", "B2", "B3", "C2", "C3"]
        assert new CellLabelIterator("A4", "C3").collect { it } == []
    }

    void testIterator_points() {
        assert new CellLabelIterator(0, 0, 2, 1).collect { it } == ["A1", "A2", "A3", "B1", "B2", "B3"]
        assert new CellLabelIterator(1, 0, 2, 1).collect { it } == ["A2", "A3", "B2", "B3"]
        assert new CellLabelIterator(1, 0, 2, 2).collect { it } == ["A2", "A3", "B2", "B3", "C2", "C3"]
        assert new CellLabelIterator(3, 0, 2, 2).collect { it } == []
    }
}
