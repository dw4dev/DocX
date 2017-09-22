using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Novacode
{
    /// <summary>
    /// 格式 - 文字框線
    /// <para>參考來源 ref source：http://officeopenxml.com/WPtextBorders.php</para>
    /// add by Davidson @ 20170922
    /// </summary>
    public class TextBorders
    {
        /// <summary>
        /// 框線樣式
        /// </summary>
        public BorderStyle borderStyle { get; set; } = BorderStyle.Tcbs_none;

        /// <summary>
        /// 框線寬度
        /// </summary>
        public BorderSize borderWidth { get; set; } = BorderSize.two;

        /// <summary>
        /// Specifies the spacing offset. Values are specified in points (1/72nd of an inch).
        /// </summary>
        public int borderSpace { get; set; } = 0;
        /// <summary>
        /// 框線顏色
        /// <para>若將顏色設定為 Transparent 則表示顏色為 auto ；暫時用此方式替代。</para>
        /// </summary>
        public Color borderColor { get; set; } = Color.Transparent;

        /// <summary>
        /// Specifies whether the border should be modified to create the appearance of a shadow. 
        /// For right and bottom borders, this is done by duplicating the border below and right of the normal location. 
        /// For the right and top borders, this is done by moving the border down and to the right of the original location. 
        /// Permitted values are true and false.
        /// </summary>
        private bool borderShadow { get; set; }

        /// <summary>
        /// Specifies whether the border should be modified to create a frame effect 
        ///     by reversing the border's appearance from the edge nearest the text to the edge furthest from the text. 
        /// Permitted values are true and false.
        /// </summary>
        public bool borderFrame { get; set; } = true;


        public TextBorders()
        {
            borderStyle = BorderStyle.Tcbs_single;
        }

        public TextBorders(BorderStyle bs, BorderSize sz, Color bc, int space = 0)
        {
            borderStyle = bs;
            borderWidth = sz;
            borderColor = bc;
            borderSpace = space;
        }
    }
}
