// BizTalk Typed BAM API Generator
// Copyright (C) 2008-Present Thomas F. Abraham. All Rights Reserved.
// Copyright (c) 2007 Darren Jefford. All Rights Reserved.
// Licensed under the MIT License. See License.txt in the project root.

using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace Shared
{
    /// <summary>
    /// Extracts the XML for a BizTalk BAM definition from an OpenXML Excel XLSX file.
    /// </summary>
    internal static class BamDefinitionXmlExporter
    {
        public static string GetBamDefinitionXml(string xlsFileName)
        {
            if (!File.Exists(xlsFileName))
            {
                throw new ArgumentException("File '" + xlsFileName + "' does not exist or is unavailable.");
            }

            if (!xlsFileName.EndsWith(".xlsx"))
            {
                throw new ArgumentException("File '" + xlsFileName + "' does not appear to be an .xlsx file.");
            }

            return GetBamDefinitionXmlFromXlsx(xlsFileName);
        }

        private static string GetBamDefinitionXmlFromXlsx(string xlsFileName)
        {
            using (var workbook = new XLWorkbook(xlsFileName))
            {
                IXLWorksheet bamWorksheet;
                var success = workbook.Worksheets.TryGetWorksheet("BamXmlHiddenSheet", out bamWorksheet);
                if (!success)
                {
                    throw new ArgumentException("ERROR: Could not find hidden BAM worksheet BamXmlHiddenSheet.");
                }
                var bamColumn = bamWorksheet.FirstColumn();
                var bamCells = bamColumn.CellsUsed();

                if (bamCells.Count() == 0)
                {
                    throw new ArgumentException(
                        "ERROR: Could not find hidden BAM worksheet or found no BAM XML on the worksheet. Expected to find BAM XML at cell BamXmlHiddenSheet!A1.");
                }
                StringBuilder sb = new StringBuilder();
                foreach (var cell in bamCells)
                {
                    sb.Append(cell.Value.ToString());
                }

                return sb.ToString();
            }
        }
    }
}
