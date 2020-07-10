package software.leonov.jcells;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import org.apache.poi.hssf.usermodel.HSSFExtendedColor;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Common {@code Color}s and utility methods.
 * 
 * @author Zhenya Leonov
 */
public final class Colors {

    public static final Color BROWN = HSSFColorPredefined.BROWN.getColor();
    public static final Color OLIVE_GREEN = HSSFColorPredefined.OLIVE_GREEN.getColor();
    public static final Color DARK_GREEN = HSSFColorPredefined.DARK_GREEN.getColor();
    public static final Color DARK_TEAL = HSSFColorPredefined.DARK_TEAL.getColor();
    public static final Color DARK_BLUE = HSSFColorPredefined.DARK_BLUE.getColor();
    public static final Color INDIGO = HSSFColorPredefined.INDIGO.getColor();
    public static final Color GREY_80_PERCENT = HSSFColorPredefined.GREY_80_PERCENT.getColor();
    public static final Color ORANGE = HSSFColorPredefined.ORANGE.getColor();
    public static final Color DARK_YELLOW = HSSFColorPredefined.DARK_YELLOW.getColor();
    public static final Color GREEN = HSSFColorPredefined.GREEN.getColor();
    public static final Color TEAL = HSSFColorPredefined.TEAL.getColor();
    public static final Color BLUE = HSSFColorPredefined.BLUE.getColor();
    public static final Color BLUE_GREY = HSSFColorPredefined.BLUE_GREY.getColor();
    public static final Color GREY_50_PERCENT = HSSFColorPredefined.GREY_50_PERCENT.getColor();
    public static final Color RED = HSSFColorPredefined.RED.getColor();
    public static final Color LIGHT_ORANGE = HSSFColorPredefined.LIGHT_ORANGE.getColor();
    public static final Color LIME = HSSFColorPredefined.LIME.getColor();
    public static final Color SEA_GREEN = HSSFColorPredefined.SEA_GREEN.getColor();
    public static final Color AQUA = HSSFColorPredefined.AQUA.getColor();
    public static final Color LIGHT_BLUE = HSSFColorPredefined.LIGHT_BLUE.getColor();
    public static final Color VIOLET = HSSFColorPredefined.VIOLET.getColor();
    public static final Color GREY_40_PERCENT = HSSFColorPredefined.GREY_40_PERCENT.getColor();
    public static final Color PINK = HSSFColorPredefined.PINK.getColor();
    public static final Color GOLD = HSSFColorPredefined.GOLD.getColor();
    public static final Color YELLOW = HSSFColorPredefined.YELLOW.getColor();
    public static final Color BRIGHT_GREEN = HSSFColorPredefined.BRIGHT_GREEN.getColor();
    public static final Color TURQUOISE = HSSFColorPredefined.TURQUOISE.getColor();
    public static final Color DARK_RED = HSSFColorPredefined.DARK_RED.getColor();
    public static final Color SKY_BLUE = HSSFColorPredefined.SKY_BLUE.getColor();
    public static final Color PLUM = HSSFColorPredefined.PLUM.getColor();
    public static final Color GREY_25_PERCENT = HSSFColorPredefined.GREY_25_PERCENT.getColor();
    public static final Color ROSE = HSSFColorPredefined.ROSE.getColor();
    public static final Color LIGHT_YELLOW = HSSFColorPredefined.LIGHT_YELLOW.getColor();
    public static final Color LIGHT_GREEN = HSSFColorPredefined.LIGHT_GREEN.getColor();
    public static final Color LIGHT_TURQUOISE = HSSFColorPredefined.LIGHT_TURQUOISE.getColor();
    public static final Color PALE_BLUE = HSSFColorPredefined.PALE_BLUE.getColor();
    public static final Color LAVENDER = HSSFColorPredefined.LAVENDER.getColor();
    public static final Color WHITE = HSSFColorPredefined.WHITE.getColor();
    public static final Color CORNFLOWER_BLUE = HSSFColorPredefined.CORNFLOWER_BLUE.getColor();
    public static final Color LEMON_CHIFFON = HSSFColorPredefined.LEMON_CHIFFON.getColor();
    public static final Color MAROON = HSSFColorPredefined.MAROON.getColor();
    public static final Color ORCHID = HSSFColorPredefined.ORCHID.getColor();
    public static final Color CORAL = HSSFColorPredefined.CORAL.getColor();
    public static final Color ROYAL_BLUE = HSSFColorPredefined.ROYAL_BLUE.getColor();
    public static final Color LIGHT_CORNFLOWER_BLUE = HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getColor();
    public static final Color TAN = HSSFColorPredefined.TAN.getColor();

    private Colors() {
    }

    public static Color rgb(final int red, final int green, final int blue) {
        checkArgument(red >= 0, "red < 0");
        checkArgument(red <= 255, "red > 255");
        checkArgument(green >= 0, "green < 0");
        checkArgument(green <= 255, "green > 255");
        checkArgument(blue >= 0, "blue < 0");
        checkArgument(blue <= 255, "blue > 255");
        return new XSSFColor(new byte[] { (byte) red, (byte) green, (byte) blue }, null);
    }

    /**
     * Returns a new {@code Color} represented by a hexadecimal string.
     * 
     * @param hex the hexadecimal string to decode
     * @return a new {@code Color} represented by a hexadecimal string
     */
    public static Color hex(final String hex) {
        checkNotNull(hex, "hex == null");
        return new XSSFColor(java.awt.Color.decode(hex), null);
    }

    static short getIndex(final Color color) {
        if (color instanceof ExtendedColor)
            return ((ExtendedColor) color).getIndex();
        else if (color instanceof HSSFColor)
            return ((HSSFColor) color).getIndex();
        else if (color instanceof HSSFExtendedColor)
            return ((HSSFExtendedColor) color).getIndex();
        else if (color instanceof XSSFColor)
            return ((XSSFColor) color).getIndex();
        throw new IllegalArgumentException("unsupported Color: " + color.getClass().getSimpleName());
    }

}
