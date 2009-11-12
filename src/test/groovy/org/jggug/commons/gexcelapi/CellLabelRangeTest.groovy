package org.jggug.commons.gexcelapi

class CellLabelRangeTest extends GroovyTestCase {

    void testRange_label() {
        assert new CellLabelRange("A1", "B3").collect { it } == ["A1", "A2", "A3", "B1", "B2", "B3"]
        assert new CellLabelRange("A2", "B3").collect { it } == ["A2", "A3", "B2", "B3"]
        assert new CellLabelRange("A2", "C3").collect { it } == ["A2", "A3", "B2", "B3", "C2", "C3"]
        assert new CellLabelRange("A4", "C3").collect { it } == []
    }

    void testRange_points() {
        assert new CellLabelRange(0, 0, 2, 1).collect { it } == ["A1", "A2", "A3", "B1", "B2", "B3"]
        assert new CellLabelRange(1, 0, 2, 1).collect { it } == ["A2", "A3", "B2", "B3"]
        assert new CellLabelRange(1, 0, 2, 2).collect { it } == ["A2", "A3", "B2", "B3", "C2", "C3"]
        assert new CellLabelRange(3, 0, 2, 2).collect { it } == []
    }

    void testInScope() {
        def range = new CellLabelRange("A1", "B2")
        assert "A1" in range
        assert "A2" in range
        assert "A3" in range == false
        assert "B1" in range
        assert "B2" in range
        assert "B3" in range == false
        assert "C1" in range == false
    }

    void testForIn() {
        def range = new CellLabelRange("A1", "B3")
        def result = []
        for (cell in range) {
            result << cell
        }
        assert result == ["A1", "A2", "A3", "B1", "B2", "B3"]
    }

}

