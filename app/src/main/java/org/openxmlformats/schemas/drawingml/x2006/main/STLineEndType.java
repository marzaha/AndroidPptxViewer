/*
 * XML Type:  ST_LineEndType
 * Namespace: http://schemas.openxmlformats.org/drawingml/2006/main
 * Java type: org.openxmlformats.schemas.drawingml.x2006.main.STLineEndType
 *
 * Automatically generated - do not modify.
 */
package org.openxmlformats.schemas.drawingml.x2006.main;


/**
 * An XML ST_LineEndType(@http://schemas.openxmlformats.org/drawingml/2006/main).
 *
 * This is an atomic type that is a restriction of org.openxmlformats.schemas.drawingml.x2006.main.STLineEndType.
 */
public interface STLineEndType extends org.apache.xmlbeans.XmlToken
{
    public static final org.apache.xmlbeans.SchemaType type = (org.apache.xmlbeans.SchemaType)
        org.apache.xmlbeans.XmlBeans.typeSystemForClassLoader(STLineEndType.class.getClassLoader(), "schemaorg_apache_xmlbeans.system.sE130CAA0A01A7CDE5A2B4FEB8B311707").resolveHandle("stlineendtype8902type");
    
    org.apache.xmlbeans.StringEnumAbstractBase enumValue();
    void set(org.apache.xmlbeans.StringEnumAbstractBase e);
    
    static final Enum NONE = Enum.forString("none");
    static final Enum TRIANGLE = Enum.forString("triangle");
    static final Enum STEALTH = Enum.forString("stealth");
    static final Enum DIAMOND = Enum.forString("diamond");
    static final Enum OVAL = Enum.forString("oval");
    static final Enum ARROW = Enum.forString("arrow");
    
    static final int INT_NONE = Enum.INT_NONE;
    static final int INT_TRIANGLE = Enum.INT_TRIANGLE;
    static final int INT_STEALTH = Enum.INT_STEALTH;
    static final int INT_DIAMOND = Enum.INT_DIAMOND;
    static final int INT_OVAL = Enum.INT_OVAL;
    static final int INT_ARROW = Enum.INT_ARROW;
    
    /**
     * Enumeration value class for org.openxmlformats.schemas.drawingml.x2006.main.STLineEndType.
     * These enum values can be used as follows:
     * <pre>
     * enum.toString(); // returns the string value of the enum
     * enum.intValue(); // returns an int value, useful for switches
     * // e.g., case Enum.INT_NONE
     * Enum.forString(s); // returns the enum value for a string
     * Enum.forInt(i); // returns the enum value for an int
     * </pre>
     * Enumeration objects are immutable singleton objects that
     * can be compared using == object equality. They have no
     * public constructor. See the constants defined within this
     * class for all the valid values.
     */
    static final class Enum extends org.apache.xmlbeans.StringEnumAbstractBase
    {
        /**
         * Returns the enum value for a string, or null if none.
         */
        public static Enum forString(java.lang.String s)
            { return (Enum)table.forString(s); }
        /**
         * Returns the enum value corresponding to an int, or null if none.
         */
        public static Enum forInt(int i)
            { return (Enum)table.forInt(i); }
        
        private Enum(java.lang.String s, int i)
            { super(s, i); }
        
        static final int INT_NONE = 1;
        static final int INT_TRIANGLE = 2;
        static final int INT_STEALTH = 3;
        static final int INT_DIAMOND = 4;
        static final int INT_OVAL = 5;
        static final int INT_ARROW = 6;
        
        public static final org.apache.xmlbeans.StringEnumAbstractBase.Table table =
            new org.apache.xmlbeans.StringEnumAbstractBase.Table
        (
            new Enum[]
            {
                new Enum("none", INT_NONE),
                new Enum("triangle", INT_TRIANGLE),
                new Enum("stealth", INT_STEALTH),
                new Enum("diamond", INT_DIAMOND),
                new Enum("oval", INT_OVAL),
                new Enum("arrow", INT_ARROW),
            }
        );
        private static final long serialVersionUID = 1L;
        private java.lang.Object readResolve() { return forInt(intValue()); } 
    }
    
    /**
     * A factory class with static methods for creating instances
     * of this type.
     */
    
}
