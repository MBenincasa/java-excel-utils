package io.github.mbenincasa.javaexcelutils.model.parser;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * Class used to define the rules to be applied for a Sheet in a list of objects
 * @author Mirko Benincasa
 * @since 0.5.0
 */
@AllArgsConstructor
@Getter
public class ExcelListParserMapping {

    /**
     * Source cell name
     */
    private String startingCell;

    /**
     * List direction
     */
    private Direction direction;

    /**
     * The distance, according to the line of direction, between two adjacent sources
     */
    private Integer jumpCells;
}
