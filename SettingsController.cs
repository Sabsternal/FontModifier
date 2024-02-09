using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FontModifier
{
    public partial class SettingsController
    {
        private bool previousBold;
        private bool currentBold;
        private bool previousItalic;
        private bool currentItalic;
        private int previousFont;
        private int currentFont;
        private System.Drawing.Color previousColor;
        private System.Drawing.Color currentColor;
        private float previousSize;
        private float currentSize;
        public SettingsController()
        {

        }
        public bool SettingsChanged()
        {
            bool isBoldChanged = this.previousBold != this.currentBold;
            bool isItalicChanged = this.previousItalic != this.currentItalic;
            bool isFontChanged = this.previousFont != this.currentFont;
            bool isColorChanged = this.previousColor != this.currentColor;
            bool isSizeCahnged = this.previousSize != this.currentSize;
            return isBoldChanged || isColorChanged || isFontChanged || isItalicChanged || isSizeCahnged;
        }

        public void SetSettings(bool currentBold, bool currentItalic, int currentFont, System.Drawing.Color currentColor, float currentSize)
        {
            this.currentBold = currentBold;
            this.previousBold = currentBold;
            this.currentItalic = currentItalic;
            this.previousItalic = currentItalic;
            this.currentFont = currentFont;
            this.previousFont = currentFont;
            this.currentColor = currentColor;
            this.previousColor = currentColor;
            this.currentSize = currentSize;
            this.previousSize = currentSize;
        }

        public void UpdateSettings(bool currentBold, bool currentItalic, int currentFont, System.Drawing.Color currentColor, float currentSize)
        {
            this.previousBold = this.currentBold;
            this.previousColor = this.currentColor;
            this.previousFont = this.currentFont;
            this.previousItalic = this.currentItalic;
            this.previousSize = this.currentSize;
            this.currentBold = currentBold;
            this.currentColor = currentColor;
            this.currentFont = currentFont;
            this.currentItalic = currentItalic;
            this.currentSize = currentSize;
        }
    }
}
