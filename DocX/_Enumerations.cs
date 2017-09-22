using System;
using System.ComponentModel;

namespace Novacode
{

    public enum ListItemType
    {
        Bulleted,
        Numbered
    }

    public enum SectionBreakType
    {
        defaultNextPage,
        evenPage,
        oddPage,
        continuous
    }


    public enum ContainerType
    {
        None,
        TOC,
        Section,
        Cell,
        Table,
        Header,
        Footer,
        Paragraph,
        Body
    }

    public enum PageNumberFormat
    {
        normal,
        roman
    }

    /// <summary>
    /// 框線寬度
    /// </summary>
    public enum BorderSize
    {
        /// <summary>
        /// 1/4 pt | val = 8 * 1/4 = 2
        /// </summary>
        one = 2,
        /// <summary>
        /// 1/2 pt | val= 8 * 1/2 = 4
        /// </summary>
        two = 4,
        /// <summary>
        /// 3/4 pt | val = 8 * 3/4 = 6
        /// </summary>
        three = 6,
        /// <summary>
        /// 1 pt | val = 8 * 1 = 8
        /// </summary>
        four = 8,
        /// <summary>
        /// 1 1/2 pt | val = 8 * 1 1/2 = 12
        /// </summary>
        five = 12,
        /// <summary>
        /// 2 1/4 pt | val = 8 * 2 1/4 = 18
        /// </summary>
        six = 18,
        /// <summary>
        /// 3 pt | val = 8 * 3 = 24
        /// </summary>
        seven = 24,
        /// <summary>
        /// 4 1/2 pt | val = 8 * 4 1/2 = 36
        /// </summary>
        eight = 36,
        /// <summary>
        /// 6 pt | val = 8 * 6 = 48
        /// </summary>
        nine = 48
    }

    

    public enum EditRestrictions
    {
        none,
        readOnly,
        forms,
        comments,
        trackedChanges
    }

    /// <summary>
    /// <para>(字元|段落|)框線樣式</para>
    /// Table Cell Border styles
    /// Added by lckuiper @ 20101117
    /// source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx
    /// </summary>
    public enum BorderStyle
    {
        /// <summary>
        /// 無框線
        /// <para>no border</para>
        /// </summary>
        Tcbs_none = 0,
        /// <summary>
        /// 實線
        /// <para>a single line</para>
        /// </summary>
        Tcbs_single,
        /// <summary>
        /// <para>a single line</para>
        /// </summary>
        Tcbs_thick,
        /// <summary>
        /// 雙實線
        /// <para>a double line</para>
        /// </summary>
        Tcbs_double,
        /// <summary>
        /// <para>a dotted line</para>
        /// </summary>
        Tcbs_dotted,
        /// <summary>
        /// 短線 ------
        /// <para>a dashed line</para>
        /// </summary>
        Tcbs_dashed,
        /// <summary>
        /// 點與短線 .-.-.
        /// <para>a line with alternating dots and dashes </para>
        /// </summary>
        Tcbs_dotDash,
        /// <summary>
        /// 點點短線 ..-..-..-
        /// <para></para>
        /// </summary>
        Tcbs_dotDotDash,
        /// <summary>
        /// <para>a triple line</para>
        /// </summary>
        Tcbs_triple,
        /// <summary>
        /// <para>a thick line contained within a thin line with a small intermediate gap</para>
        /// </summary>
        Tcbs_thinThickSmallGap,
        /// <summary>
        /// <para>a thick line contained within a thin line with a small intermediate gap</para>
        /// </summary>
        Tcbs_thickThinSmallGap,
        /// <summary>
        /// <para>a thin-thick-thin line with a small gap</para>
        /// </summary>
        Tcbs_thinThickThinSmallGap,
        /// <summary>
        /// <para>a thick line contained within a thin line with a medium-sized intermediate gap</para>
        /// </summary>
        Tcbs_thinThickMediumGap,
        /// <summary>
        /// <para>a thick line contained within a thin line with a medium-sized intermediate gap</para>
        /// </summary>
        Tcbs_thickThinMediumGap,
        /// <summary>
        /// <para>a thin-thick-thin line with a medium gap</para>
        /// </summary>
        Tcbs_thinThickThinMediumGap,
        /// <summary>
        /// <para>a thick line contained within a thin line with a medium-sized intermediate gap</para>
        /// </summary>
        Tcbs_thinThickLargeGap,
        /// <summary>
        /// <para>a thick line contained within a thin line with a large-sized intermediate gap</para>
        /// </summary>
        Tcbs_thickThinLargeGap,
        /// <summary>
        /// <para>a thin-thick-thin line with a large gap</para>
        /// </summary>
        Tcbs_thinThickThinLargeGap,
        /// <summary>
        /// 波浪線
        /// <para>a wavy line</para>
        /// </summary>
        Tcbs_wave,
        /// <summary>
        /// 雙波浪線
        /// <para>a double wavy line</para>
        /// </summary>
        Tcbs_doubleWave,
        /// <summary>
        /// 小間隔短線 - - -  
        /// <para>a dashed line with small gaps</para>
        /// </summary>
        Tcbs_dashSmallGap,
        /// <summary>
        /// 點線 ......
        /// <para>a line with a series of alternating thin and thick strokes </para>
        /// </summary>
        Tcbs_dashDotStroked,
        /// <summary>
        /// <para> a three-staged gradient line, getting darker towards the paragraph</para>
        /// </summary>
        Tcbs_threeDEmboss,
        /// <summary>
        /// <para>a three-staged gradient like, getting darker away from the paragraph</para>
        /// </summary>
        Tcbs_threeDEngrave,
        /// <summary>
        /// an outset set of lines
        /// </summary>
        Tcbs_outset,
        /// <summary>
        /// an inset set of lines
        /// </summary>
        Tcbs_inset,
        /// <summary>
        /// 無框線
        /// <para>no border</para>
        /// </summary>
        Tcbs_nil
    }

    /// <summary>
    /// Table Cell Border Types
    /// Added by lckuiper @ 20101117
    /// source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx
    /// </summary>
    public enum TableCellBorderType
    {
        Top,
        Bottom,
        Left,
        Right,
        InsideH,
        InsideV,
        TopLeftToBottomRight,
        TopRightToBottomLeft
    }

    /// <summary>
    /// Table Border Types
    /// Added by lckuiper @ 20101117
    /// source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tableborders.aspx
    /// </summary>
    public enum TableBorderType
    {
        Top,
        Bottom,
        Left,
        Right,
        InsideH,
        InsideV
    }

    // Patch 7398 added by lckuiper on Nov 16th 2010 @ 2:23 PM
    public enum VerticalAlignment
    {
        Top,
        Center,
        Bottom
    };

    public enum Orientation
    {
        Portrait,
        Landscape
    };

    public enum XmlDocument
    {
        Main,
        HeaderOdd,
        HeaderEven,
        HeaderFirst,
        FooterOdd,
        FooterEven,
        FooterFirst
    };

    public enum MatchFormattingOptions
    {
        ExactMatch,
        SubsetMatch
    };

    public enum Script
    {
        superscript,
        subscript,
        none
    }
    /// <summary>
    /// 文字醒目提示色彩
    /// </summary>
    public enum Highlight
    {
        /// <summary>
        /// 黃色
        /// </summary>
        yellow,
        /// <summary>
        /// 亮綠色
        /// </summary>
        green,
        /// <summary>
        /// 淺粉藍
        /// </summary>
        cyan,
        /// <summary>
        /// 粉紅色
        /// </summary>
        magenta,
        /// <summary>
        /// 藍色
        /// </summary>
        blue,
        /// <summary>
        /// 紅色
        /// </summary>
        red,
        /// <summary>
        /// 深藍色
        /// </summary>
        darkBlue,
        /// <summary>
        /// 藍綠色
        /// </summary>
        darkCyan,
        /// <summary>
        /// 綠色
        /// </summary>
        darkGreen,
        /// <summary>
        /// 紫羅蘭色
        /// </summary>
        darkMagenta,
        /// <summary>
        /// 深紅色
        /// </summary>
        darkRed,
        /// <summary>
        /// 深黃色
        /// </summary>
        darkYellow,
        /// <summary>
        /// 灰50%
        /// </summary>
        darkGray,
        /// <summary>
        /// 灰25%
        /// </summary>
        lightGray,
        /// <summary>
        /// 黑色
        /// </summary>
        black,
        /// <summary>
        /// 無色彩
        /// </summary>
        none
    };

    public enum UnderlineStyle
    {
        none = 0,
        singleLine = 1,
        words = 2,
        doubleLine = 3,
        dotted = 4,
        thick = 6,
        dash = 7,
        dotDash = 9,
        dotDotDash = 10,
        wave = 11,
        dottedHeavy = 20,
        dashedHeavy = 23,
        dashDotHeavy = 25,
        dashDotDotHeavy = 26,
        dashLongHeavy = 27,
        dashLong = 39,
        wavyDouble = 43,
        wavyHeavy = 55,
    };

    public enum StrikeThrough
    {
        none,
        strike,
        doubleStrike
    };

    public enum Misc
    {
        none,
        shadow,
        outline,
        outlineShadow,
        emboss,
        engrave
    };

    /// <summary>
    /// Change the caps style of text, for use with Append and AppendLine.
    /// </summary>
    public enum CapsStyle
    {
        /// <summary>
        /// No caps, make all characters are lowercase.
        /// </summary>
        none,

        /// <summary>
        /// All caps, make every character uppercase.
        /// </summary>
        caps,

        /// <summary>
        /// Small caps, make all characters capital but with a small font size.
        /// </summary>
        smallCaps
    };

    /// <summary>
    /// Designs\Styles that can be applied to a table.
    /// </summary>
    public enum TableDesign
    {
        Custom,
        TableNormal,
        TableGrid,
        LightShading,
        LightShadingAccent1,
        LightShadingAccent2,
        LightShadingAccent3,
        LightShadingAccent4,
        LightShadingAccent5,
        LightShadingAccent6,
        LightList,
        LightListAccent1,
        LightListAccent2,
        LightListAccent3,
        LightListAccent4,
        LightListAccent5,
        LightListAccent6,
        LightGrid,
        LightGridAccent1,
        LightGridAccent2,
        LightGridAccent3,
        LightGridAccent4,
        LightGridAccent5,
        LightGridAccent6,
        MediumShading1,
        MediumShading1Accent1,
        MediumShading1Accent2,
        MediumShading1Accent3,
        MediumShading1Accent4,
        MediumShading1Accent5,
        MediumShading1Accent6,
        MediumShading2,
        MediumShading2Accent1,
        MediumShading2Accent2,
        MediumShading2Accent3,
        MediumShading2Accent4,
        MediumShading2Accent5,
        MediumShading2Accent6,
        MediumList1,
        MediumList1Accent1,
        MediumList1Accent2,
        MediumList1Accent3,
        MediumList1Accent4,
        MediumList1Accent5,
        MediumList1Accent6,
        MediumList2,
        MediumList2Accent1,
        MediumList2Accent2,
        MediumList2Accent3,
        MediumList2Accent4,
        MediumList2Accent5,
        MediumList2Accent6,
        MediumGrid1,
        MediumGrid1Accent1,
        MediumGrid1Accent2,
        MediumGrid1Accent3,
        MediumGrid1Accent4,
        MediumGrid1Accent5,
        MediumGrid1Accent6,
        MediumGrid2,
        MediumGrid2Accent1,
        MediumGrid2Accent2,
        MediumGrid2Accent3,
        MediumGrid2Accent4,
        MediumGrid2Accent5,
        MediumGrid2Accent6,
        MediumGrid3,
        MediumGrid3Accent1,
        MediumGrid3Accent2,
        MediumGrid3Accent3,
        MediumGrid3Accent4,
        MediumGrid3Accent5,
        MediumGrid3Accent6,
        DarkList,
        DarkListAccent1,
        DarkListAccent2,
        DarkListAccent3,
        DarkListAccent4,
        DarkListAccent5,
        DarkListAccent6,
        ColorfulShading,
        ColorfulShadingAccent1,
        ColorfulShadingAccent2,
        ColorfulShadingAccent3,
        ColorfulShadingAccent4,
        ColorfulShadingAccent5,
        ColorfulShadingAccent6,
        ColorfulList,
        ColorfulListAccent1,
        ColorfulListAccent2,
        ColorfulListAccent3,
        ColorfulListAccent4,
        ColorfulListAccent5,
        ColorfulListAccent6,
        ColorfulGrid,
        ColorfulGridAccent1,
        ColorfulGridAccent2,
        ColorfulGridAccent3,
        ColorfulGridAccent4,
        ColorfulGridAccent5,
        ColorfulGridAccent6,
        None
    };

    /// <summary>
    /// How a Table should auto resize.
    /// </summary>
    public enum AutoFit
    {
        /// <summary>
        /// Autofit to Table contents.
        /// </summary>
        Contents,

        /// <summary>
        /// Autofit to Window.
        /// </summary>
        Window,

        /// <summary>
        /// Autofit to Column width.
        /// </summary>
        ColumnWidth,
        /// <summary>
        ///  Autofit to Fixed column width
        /// </summary>
        Fixed
    };

    public enum RectangleShapes
    {
        rect,
        roundRect,
        snip1Rect,
        snip2SameRect,
        snip2DiagRect,
        snipRoundRect,
        round1Rect,
        round2SameRect,
        round2DiagRect
    };

    public enum BasicShapes
    {
        ellipse,
        triangle,
        rtTriangle,
        parallelogram,
        trapezoid,
        diamond,
        pentagon,
        hexagon,
        heptagon,
        octagon,
        decagon,
        dodecagon,
        pie,
        chord,
        teardrop,
        frame,
        halfFrame,
        corner,
        diagStripe,
        plus,
        plaque,
        can,
        cube,
        bevel,
        donut,
        noSmoking,
        blockArc,
        foldedCorner,
        smileyFace,
        heart,
        lightningBolt,
        sun,
        moon,
        cloud,
        arc,
        backetPair,
        bracePair,
        leftBracket,
        rightBracket,
        leftBrace,
        rightBrace
    };

    public enum BlockArrowShapes
    {
        rightArrow,
        leftArrow,
        upArrow,
        downArrow,
        leftRightArrow,
        upDownArrow,
        quadArrow,
        leftRightUpArrow,
        bentArrow,
        uturnArrow,
        leftUpArrow,
        bentUpArrow,
        curvedRightArrow,
        curvedLeftArrow,
        curvedUpArrow,
        curvedDownArrow,
        stripedRightArrow,
        notchedRightArrow,
        homePlate,
        chevron,
        rightArrowCallout,
        downArrowCallout,
        leftArrowCallout,
        upArrowCallout,
        leftRightArrowCallout,
        quadArrowCallout,
        circularArrow
    };

    public enum EquationShapes
    {
        mathPlus,
        mathMinus,
        mathMultiply,
        mathDivide,
        mathEqual,
        mathNotEqual
    };

    public enum FlowchartShapes
    {
        flowChartProcess,
        flowChartAlternateProcess,
        flowChartDecision,
        flowChartInputOutput,
        flowChartPredefinedProcess,
        flowChartInternalStorage,
        flowChartDocument,
        flowChartMultidocument,
        flowChartTerminator,
        flowChartPreparation,
        flowChartManualInput,
        flowChartManualOperation,
        flowChartConnector,
        flowChartOffpageConnector,
        flowChartPunchedCard,
        flowChartPunchedTape,
        flowChartSummingJunction,
        flowChartOr,
        flowChartCollate,
        flowChartSort,
        flowChartExtract,
        flowChartMerge,
        flowChartOnlineStorage,
        flowChartDelay,
        flowChartMagneticTape,
        flowChartMagneticDisk,
        flowChartMagneticDrum,
        flowChartDisplay
    };

    public enum StarAndBannerShapes
    {
        irregularSeal1,
        irregularSeal2,
        star4,
        star5,
        star6,
        star7,
        star8,
        star10,
        star12,
        star16,
        star24,
        star32,
        ribbon,
        ribbon2,
        ellipseRibbon,
        ellipseRibbon2,
        verticalScroll,
        horizontalScroll,
        wave,
        doubleWave
    };

    public enum CalloutShapes
    {
        wedgeRectCallout,
        wedgeRoundRectCallout,
        wedgeEllipseCallout,
        cloudCallout,
        borderCallout1,
        borderCallout2,
        borderCallout3,
        accentCallout1,
        accentCallout2,
        accentCallout3,
        callout1,
        callout2,
        callout3,
        accentBorderCallout1,
        accentBorderCallout2,
        accentBorderCallout3
    };

    /// <summary>
    /// Text alignment of a Paragraph.
    /// </summary>
    public enum Alignment
    {
        /// <summary>
        /// Align Paragraph to the left.
        /// </summary>
        left,

        /// <summary>
        /// Align Paragraph as centered.
        /// </summary>
        center,

        /// <summary>
        /// Align Paragraph to the right.
        /// </summary>
        right,

        /// <summary>
        /// (Justified) Align Paragraph to both the left and right margins, adding extra space between content as necessary.
        /// </summary>
        both
    };

    public enum Direction
    {
        LeftToRight,
        RightToLeft
    };

    /// <summary>
    /// Paragraph edit types
    /// </summary>
    internal enum EditType
    {
        /// <summary>
        /// A ins is a tracked insertion
        /// </summary>
        ins,

        /// <summary>
        /// A del is  tracked deletion
        /// </summary>
        del
    }

    /// <summary>
    /// Custom property types.
    /// </summary>
    internal enum CustomPropertyType
    {
        /// <summary>
        /// System.String
        /// </summary>
        Text,

        /// <summary>
        /// System.DateTime
        /// </summary>
        Date,

        /// <summary>
        /// System.Int32
        /// </summary>
        NumberInteger,

        /// <summary>
        /// System.Double
        /// </summary>
        NumberDecimal,

        /// <summary>
        /// System.Boolean
        /// </summary>
        YesOrNo
    }

    /// <summary>
    /// Text types in a Run
    /// </summary>
    public enum RunTextType
    {
        /// <summary>
        /// System.String
        /// </summary>
        Text,

        /// <summary>
        /// System.String
        /// </summary>
        DelText,
    }
    public enum LineSpacingType
    {
        Line,
        Before,
        After
    }

    public enum LineSpacingTypeAuto
    {
        AutoBefore,
        AutoAfter,
        Auto,
        None
    }

    /// <summary>
    /// Cell margin for all sides of the table cell.
    /// </summary>
    public enum TableCellMarginType
    {
        /// <summary>
        /// The left cell margin.
        /// </summary>
        left,
        /// <summary>
        /// The right cell margin.
        /// </summary>
        right,
        /// <summary>
        /// The bottom cell margin.
        /// </summary>
        bottom,
        /// <summary>
        /// The top cell margin.
        /// </summary>
        top
    }

    public enum HeadingType
    {
        [Description("Heading1")]
        Heading1,

        [Description("Heading2")]
        Heading2,

        [Description("Heading3")]
        Heading3,

        [Description("Heading4")]
        Heading4,

        [Description("Heading5")]
        Heading5,

        [Description("Heading6")]
        Heading6,

        [Description("Heading7")]
        Heading7,

        [Description("Heading8")]
        Heading8,

        [Description("Heading9")]
        Heading9,


        /*		 
         * The Character Based Headings below do not work in the same way as the headings 1-9 above, but appear on the same list in word. 
         * I have kept them here for reference in case somebody else things its just a matter of adding them in to gain extra headings
         */
        #region Other character (NOT paragraph) based Headings
        //[Description("NoSpacing")]
        //NoSpacing,

        //[Description("Title")]
        //Title,

        //[Description("Subtitle")]
        //Subtitle,

        //[Description("Quote")]
        //Quote,

        //[Description("IntenseQuote")]
        //IntenseQuote,

        //[Description("Emphasis")]
        //Emphasis,

        //[Description("IntenseEmphasis")]
        //IntenseEmphasis,

        //[Description("Strong")]
        //Strong,

        //[Description("ListParagraph")]
        //ListParagraph,

        //[Description("SubtleReference")]
        //SubtleReference,

        //[Description("IntenseReference")]
        //IntenseReference,

        //[Description("BookTitle")]
        //BookTitle, 
        #endregion


    }
    public enum TextDirection
    {
        btLr,
        right
    };

    /// <summary>
    /// Represents the switches set on a TOC.
    /// </summary>
    [Flags]
    public enum TableOfContentsSwitches
    {
        None = 0 << 0,
        [Description("\\a")]
        A = 1 << 0,
        [Description("\\b")]
        B = 1 << 1,
        [Description("\\c")]
        C = 1 << 2,
        [Description("\\d")]
        D = 1 << 3,
        [Description("\\f")]
        F = 1 << 4,
        [Description("\\h")]
        H = 1 << 5,
        [Description("\\l")]
        L = 1 << 6,
        [Description("\\n")]
        N = 1 << 7,
        [Description("\\o")]
        O = 1 << 8,
        [Description("\\p")]
        P = 1 << 9,
        [Description("\\s")]
        S = 1 << 10,
        [Description("\\t")]
        T = 1 << 11,
        [Description("\\u")]
        U = 1 << 12,
        [Description("\\w")]
        W = 1 << 13,
        [Description("\\x")]
        X = 1 << 14,
        [Description("\\z")]
        Z = 1 << 15,
    }

}