/////////////////////////////////////////////////////////////////////////////
//
// Project ProjectForge Community Edition
//         www.projectforge.org
//
// Copyright (C) 2001-2013 Kai Reinhard (k.reinhard@micromata.de)
//
// ProjectForge is dual-licensed.
//
// This community edition is free software; you can redistribute it and/or
// modify it under the terms of the GNU General Public License as published
// by the Free Software Foundation; version 3 of the License.
//
// This community edition is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
// Public License for more details.
//
// You should have received a copy of the GNU General Public License along
// with this program; if not, see http://www.gnu.org/licenses/.
//
/////////////////////////////////////////////////////////////////////////////

package org.projectforge.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;

import org.junit.Test;

public class ExportWorkbookTest
{
  private static final org.apache.log4j.Logger log = org.apache.log4j.Logger.getLogger(ExportWorkbookTest.class);

  @Test
  public void exportExcel() throws IOException
  {
    final ExportWorkbook workbook = new ExportWorkbook();
    final ExportSheet sheet = workbook.addSheet("Test");
    sheet.getContentProvider().setColWidths(20, 20, 20);
    sheet.addRow().setValues("Type", "Precision", "result");
    sheet.addRow().setValues("Java output", ".", "Tue Sep 28 00:27:10 UTC 2010");
    sheet.addRow().setValues("int", "-", 1234);
    sheet.addRow().setValues("BigDecimal", "-", new BigDecimal("123123123.123123123123"));
    final File file = new File("target/test-excel.xls");
    log.info("Writing Excel test sheet to work directory: " + file.getAbsolutePath());
    workbook.write(new FileOutputStream(file));
  }
}
