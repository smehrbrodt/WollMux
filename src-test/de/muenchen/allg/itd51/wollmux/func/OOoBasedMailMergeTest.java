package de.muenchen.allg.itd51.wollmux.func;

import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Set;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.sdbc.SQLException;

import junit.framework.TestCase;
import de.muenchen.allg.itd51.wollmux.func.OOoBasedMailMerge.MemoryDataSource;
import de.muenchen.allg.itd51.wollmux.func.OOoBasedMailMerge.MyColumn;
import de.muenchen.allg.itd51.wollmux.func.OOoBasedMailMerge.MyRow;

public class OOoBasedMailMergeTest extends TestCase
{

  public void testMemoryDataSource() throws Exception
  {
    MyRow row = null;
    Set<String> names = null;

    MemoryDataSource mds = new MemoryDataSource();
    HashMap<String, String> dataset = new HashMap<String, String>();

    dataset.put("#DS", "1");
    dataset.put("Name", "Name 1");
    mds.addDataset(dataset);

    dataset.put("#DS", "2");
    dataset.put("Name", "Name 2");
    mds.addDataset(dataset);

    // current row it beforeFirst
    assertEquals(0, mds.getRow());
    assertEquals(2, mds.getColumns().getElementNames().length);

    // test 1st row
    mds.absolute(1);
    assertEquals(1, mds.getRow());
    row = (MyRow) mds.getColumns();
    assertEquals(2, row.getElementNames().length);
    assertColumnValue("Name 1", mds, "Name");
    assertColumnValue("1", mds, "#DS");
    names = new HashSet<String>(Arrays.asList(row.getElementNames()));
    assertEquals(false, names.contains("something-different"));
    try
    {
      assertEquals("", ((MyColumn) row.getByName("something-different")).getString());
      fail();
    }
    catch (NoSuchElementException e)
    {}

    // test 2nd row
    mds.absolute(2);
    assertEquals(2, mds.getRow());
    row = (MyRow) mds.getColumns();
    assertEquals(2, row.getElementNames().length);
    assertColumnValue("Name 2", mds, "Name");
    assertColumnValue("2", mds, "#DS");
    names = new HashSet<String>(Arrays.asList(row.getElementNames()));
    assertEquals(false, names.contains("something-different"));
    try
    {
      assertEquals("", ((MyColumn) row.getByName("something-different")).getString());
      fail();
    }
    catch (NoSuchElementException e)
    {}
  }

  public static void assertColumnValue(String expected, MemoryDataSource mds,
      String column) throws SQLException, NoSuchElementException,
      WrappedTargetException
  {
    MyRow row = (MyRow) mds.getColumns();
    HashSet<String> names =
      new HashSet<String>(Arrays.asList(row.getElementNames()));
    assertEquals(true, names.contains(column));
    assertEquals(expected, ((MyColumn) row.getByName(column)).getString());
  }

  public void testMemoryDataSourceWithSelection() throws Exception
  {
    MyRow row = null;
    Set<String> names = null;

    MemoryDataSource mds = new MemoryDataSource();
    HashMap<String, String> dataset = new HashMap<String, String>();

    dataset.put("#DS", "1");
    dataset.put("Name", "Name 1");
    mds.addDataset(dataset);

    dataset.put("#DS", "2");
    dataset.put("Name", "Name 2");
    mds.addDataset(dataset);

    dataset.put("#DS", "3");
    dataset.put("Name", "Name 3");
    mds.addDataset(dataset);

    dataset.put("#DS", "4");
    dataset.put("Name", "Name 4");
    mds.addDataset(dataset);

    mds.setSelection(0, 2);
    assertEquals(2, mds.getSelectionSize());

    // current row it beforeFirst
    assertEquals(0, mds.getRow());
    assertEquals(2, mds.getColumns().getElementNames().length);

    // test 1st row
    mds.absolute(1);
    assertEquals(1, mds.getRow());
    row = (MyRow) mds.getColumns();
    assertEquals(2, row.getElementNames().length);
    names = new HashSet<String>(Arrays.asList(row.getElementNames()));
    assertEquals(true, names.contains("#DS"));
    assertEquals(true, names.contains("Name"));
    assertEquals(false, names.contains("something-different"));
    assertEquals("Name 1", ((MyColumn) row.getByName("Name")).getString());
    assertEquals("1", ((MyColumn) row.getByName("#DS")).getString());
    try
    {
      assertEquals("", ((MyColumn) row.getByName("something-different")).getString());
      fail();
    }
    catch (NoSuchElementException e)
    {}

    // test 2nd row
    mds.absolute(2);
    assertEquals(2, mds.getRow());
    row = (MyRow) mds.getColumns();
    assertEquals(2, row.getElementNames().length);
    names = new HashSet<String>(Arrays.asList(row.getElementNames()));
    assertEquals(true, names.contains("#DS"));
    assertEquals(true, names.contains("Name"));
    assertEquals(false, names.contains("something-different"));
    assertEquals("Name 2", ((MyColumn) row.getByName("Name")).getString());
    assertEquals("2", ((MyColumn) row.getByName("#DS")).getString());
    try
    {
      assertEquals("", ((MyColumn) row.getByName("something-different")).getString());
      fail();
    }
    catch (NoSuchElementException e)
    {}
  }

  public void testMemoryDataSourceWithSelectionIterate() throws Exception
  {
    MemoryDataSource mds = new MemoryDataSource();
    HashMap<String, String> dataset = new HashMap<String, String>();

    dataset.put("#DS", "1");
    dataset.put("Name", "Name 1");
    mds.addDataset(dataset);

    dataset.put("#DS", "2");
    dataset.put("Name", "Name 2");
    mds.addDataset(dataset);

    dataset.put("#DS", "3");
    dataset.put("Name", "Name 3");
    mds.addDataset(dataset);

    dataset.put("#DS", "4");
    dataset.put("Name", "Name 4");
    mds.addDataset(dataset);

    mds.setSelection(0, 2);

    mds.beforeFirst();
    assertEquals(0, mds.getRow());
    assertColumnValue("", mds, "Name");
    assertColumnValue("", mds, "#DS");

    assertEquals(true, mds.first());
    assertEquals(1, mds.getRow());
    assertColumnValue("Name 1", mds, "Name");
    assertColumnValue("1", mds, "#DS");


    assertEquals(true, mds.next());
    assertEquals(2, mds.getRow());
    assertColumnValue("Name 2", mds, "Name");
    assertColumnValue("2", mds, "#DS");

    assertEquals(false, mds.next());
    assertEquals(3, mds.getRow());
    assertColumnValue("", mds, "Name");
    assertColumnValue("", mds, "#DS");

    assertEquals(true, mds.last());
    assertEquals(2, mds.getRow());
    assertColumnValue("Name 2", mds, "Name");
    assertColumnValue("2", mds, "#DS");

    mds.afterLast();
    assertEquals(3, mds.getRow());
    assertColumnValue("", mds, "Name");
    assertColumnValue("", mds, "#DS");

    mds.setSelection(2, 4);

    mds.beforeFirst();
    assertEquals(0, mds.getRow());
    assertColumnValue("", mds, "Name");
    assertColumnValue("", mds, "#DS");

    assertEquals(false, mds.previous());
    assertEquals(0, mds.getRow());
    assertColumnValue("", mds, "Name");
    assertColumnValue("", mds, "#DS");

    assertEquals(true, mds.first());
    assertEquals(1, mds.getRow());
    assertColumnValue("Name 3", mds, "Name");
    assertColumnValue("3", mds, "#DS");

    assertEquals(true, mds.next());
    assertEquals(2, mds.getRow());
    assertColumnValue("Name 4", mds, "Name");
    assertColumnValue("4", mds, "#DS");

    assertEquals(false, mds.next());
    assertEquals(3, mds.getRow());
    assertColumnValue("", mds, "Name");
    assertColumnValue("", mds, "#DS");

    assertEquals(false, mds.next());
    assertEquals(3, mds.getRow());
    assertColumnValue("", mds, "Name");
    assertColumnValue("", mds, "#DS");

    assertEquals(true, mds.last());
    assertEquals(2, mds.getRow());
    assertColumnValue("Name 4", mds, "Name");
    assertColumnValue("4", mds, "#DS");

    mds.afterLast();
    assertEquals(3, mds.getRow());
    assertColumnValue("", mds, "Name");
    assertColumnValue("", mds, "#DS");
  }

  public void testMemoryDataSourceWithSelectionIterateEmptyDataset()
      throws Exception
  {
    MemoryDataSource mds = new MemoryDataSource();

    mds.setSelection(0, 2);
    assertEquals(0, mds.getSelectionSize());

    assertEquals(0, mds.getRow());
    assertEquals(0, mds.getColumns().getElementNames().length);

    mds.beforeFirst();
    assertEquals(0, mds.getRow());

    assertEquals(false, mds.first());
    assertEquals(1, mds.getRow());

    assertEquals(false, mds.next());
    assertEquals(1, mds.getRow());

    assertEquals(false, mds.next());
    assertEquals(1, mds.getRow());

    assertEquals(false, mds.last());
    assertEquals(0, mds.getRow());

    mds.afterLast();
    assertEquals(1, mds.getRow());
  }

}
