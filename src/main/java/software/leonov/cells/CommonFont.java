package software.leonov.cells;

import org.apache.poi.ss.usermodel.Font;

/**
 * Cross platform fonts that should work on all versions of Microsoft Excel.
 * 
 * @author Zheny Leonov
 */
public enum CommonFont {

    /**
     * Arial font.
     */
    ARIAL("Arial"),

    /**
     * Comic Sans MS font.
     */
    COMIC_SANS_MS("Comic Sans MS"),

    /**
     * Courier New font. This fond is monospaced.
     */
    COURIER_NEW("Courier New"),

    /**
     * Georgia font.
     */
    GEORGIA("Georgia"),

    /**
     * Times New Roman font.
     */
    TIMES_NEW_ROMAN("Times New Roman"),

    /**
     * Verdana font.
     */
    VERDANA("Verdana");

    private final String name;

    private CommonFont(final String name) {
        this.name = name;
    }

    /**
     * Returns the value of this {@code CommonFont} compatible with {@link Font#setFontName(String)}.
     * 
     * @return the value of this {@code CommonFont} compatible with {@link Font#setFontName(String)}.
     */
    public String getFontName() {
        return name;
    }

}
