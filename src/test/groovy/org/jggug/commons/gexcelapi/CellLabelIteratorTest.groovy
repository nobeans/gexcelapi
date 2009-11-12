package org.jggug.commons.gexcelapi

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
