/*
 * Dateiname: OOoBasedMailMerge.java
 * Projekt  : WollMux
 * Funktion : Seriendruck über den OOo MailMergeService
 * 
 * Copyright (c) 2011-2015 Landeshauptstadt München
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the European Union Public Licence (EUPL), 
 * version 1.0 (or any later version).
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * European Union Public Licence for more details.
 *
 * You should have received a copy of the European Union Public Licence
 * along with this program. If not, see 
 * http://ec.europa.eu/idabc/en/document/7330
 *
 * Änderungshistorie:
 * Datum      | Wer | Änderungsgrund
 * -------------------------------------------------------------------
 * 15.06.2011 | LUT | Erstellung
 * 12.07.2013 | JGM | Anpassungen an die neue UNO API zum setzten des Druckers
 * -------------------------------------------------------------------
 *
 * @author Christoph Lutz (D-III-ITD-D101)
 * 
 */
package de.muenchen.allg.itd51.wollmux.func;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Random;
import java.util.Vector;

import com.sun.star.beans.NamedValue;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertyChangeListener;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.beans.XVetoableChangeListener;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XNameAccess;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.io.XInputStream;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XSingleServiceFactory;
import com.sun.star.sdb.CommandType;
import com.sun.star.sdb.XColumn;
import com.sun.star.sdbc.DataType;
import com.sun.star.sdbc.SQLException;
import com.sun.star.sdbc.XArray;
import com.sun.star.sdbc.XBlob;
import com.sun.star.sdbc.XClob;
import com.sun.star.sdbc.XRef;
import com.sun.star.sdbc.XResultSet;
import com.sun.star.sdbcx.XColumnsSupplier;
import com.sun.star.sdbcx.XRowLocate;
import com.sun.star.task.XJob;
import com.sun.star.text.MailMergeEvent;
import com.sun.star.text.MailMergeType;
import com.sun.star.text.XDependentTextField;
import com.sun.star.text.XMailMergeBroadcaster;
import com.sun.star.text.XMailMergeListener;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextFramesSupplier;
import com.sun.star.text.XTextSection;
import com.sun.star.text.XTextSectionsSupplier;
import com.sun.star.text.XTextTablesSupplier;
import com.sun.star.uno.Any;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XNamingService;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.Date;
import com.sun.star.util.DateTime;
import com.sun.star.util.Time;
import com.sun.star.util.XCancellable;
import com.sun.star.view.XPrintable;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoProps;
import de.muenchen.allg.itd51.wollmux.DocumentCommand;
import de.muenchen.allg.itd51.wollmux.DocumentCommand.InsertFormValue;
import de.muenchen.allg.itd51.wollmux.DocumentCommands;
import de.muenchen.allg.itd51.wollmux.FormFieldFactory;
import de.muenchen.allg.itd51.wollmux.FormFieldFactory.FormField;
import de.muenchen.allg.itd51.wollmux.FormFieldFactory.FormFieldType;
import de.muenchen.allg.itd51.wollmux.L;
import de.muenchen.allg.itd51.wollmux.Logger;
import de.muenchen.allg.itd51.wollmux.PersistentData;
import de.muenchen.allg.itd51.wollmux.PersistentDataContainer;
import de.muenchen.allg.itd51.wollmux.PersistentDataContainer.DataID;
import de.muenchen.allg.itd51.wollmux.PrintModels;
import de.muenchen.allg.itd51.wollmux.SachleitendeVerfuegung;
import de.muenchen.allg.itd51.wollmux.SimulationResults;
import de.muenchen.allg.itd51.wollmux.SimulationResults.SimulationResultsProcessor;
import de.muenchen.allg.itd51.wollmux.TextDocumentModel;
import de.muenchen.allg.itd51.wollmux.WollMuxSingleton;
import de.muenchen.allg.itd51.wollmux.Workarounds;
import de.muenchen.allg.itd51.wollmux.XPrintModel;
import de.muenchen.allg.itd51.wollmux.dialog.mailmerge.MailMergeNew;

public class OOoBasedMailMerge
{
  private static final String DATA_SOURCE_NAME = "mmdb";

  private static final String SEP = ":";

  private static final String COLUMN_PREFIX_SINGLE_PARAMETER_FUNCTION = "WM:SP";

  private static final String COLUMN_PREFIX_CHECKBOX_FUNCTION = "WM:CB";

  private static final String COLUMN_PREFIX_MULTI_PARAMETER_FUNCTION = "WM:MP";

  private static final String TEMP_WOLLMUX_MAILMERGE_PREFIX = "WollMuxMailMerge";

  private static final String DATASOURCE_ODB_FILENAME = "datasource.odb";

  private static final String TABLE_NAME = "data";

  private static final char OPENSYMBOL_CHECKED = 0xE4C4;

  private static final char OPENSYMBOL_UNCHECKED = 0xE470;

  /**
   * Druckfunktion für den Seriendruck in ein Gesamtdokument mit Hilfe des
   * OpenOffice.org-Seriendrucks.
   * 
   * @author Christoph Lutz (D-III-ITD 5.1)
   */
  public static void oooMailMerge(final XPrintModel pmod, OutputType type)
  {
    PrintModels.setStage(pmod, L.m("Seriendruck vorbereiten"));

    File tmpDir = createMailMergeTempdir();

    // Datenquelle mit über mailMergeNewSetFormValue simulierten Daten erstellen
    MemoryDataSource ds = new MemoryDataSource();
    try
    {
      MailMergeNew.mailMergeNewSetFormValue(pmod, ds);
      if (pmod.isCanceled()) return;
    }
    catch (Exception e)
    {
      // TODO: sollen wir das wirklich machen...
      PersistentDataContainer lCont =
        PersistentData.createPersistentDataContainer(pmod.getTextDocument());
      lCont.removeData(PersistentDataContainer.DataID.FORMULARWERTE);
      Logger.debug(L.m("Formularwerte wurden aus %1 gelöscht.",
        pmod.getTextDocument().getURL()));
      // TODO: ...?
      Logger.error(
        L.m("OOo-Based-MailMerge: kann Simulationsdatenquelle nicht erzeugen!"), e);
      return;
    }
    if (ds.getSize() == 0)
    {
      WollMuxSingleton.showInfoModal(
        L.m("WollMux-Seriendruck"),
        L.m("Der Seriendruck wurde abgebrochen, da Ihr Druckauftrag keine Datensätze enthält."));
      pmod.cancel();
      return;
    }

    DocStatistics docStat = new DocStatistics();
    File inputFile =
      createAndAdjustInputFile(tmpDir, pmod.getTextDocument(), docStat);

    // Stelle fest, ob der Seriendruck in verschiedene Druckaufträge aufgeteilt
    // werden muss
    Logger.debug2(L.m("Dokumentstatistik: %1", docStat.toString()));
    int maxCriticalElements = 0;
    if (docStat.getContainedPageStyles() > maxCriticalElements)
      maxCriticalElements = docStat.getContainedPageStyles();
    if (docStat.getContainedSections() > maxCriticalElements)
      maxCriticalElements = docStat.getContainedSections();
    if (docStat.getContainedTables() > maxCriticalElements)
      maxCriticalElements = docStat.getContainedTables();
    if (docStat.getContainedTextframes() > maxCriticalElements)
      maxCriticalElements = docStat.getContainedTextframes();
    int maxProcessableDatasets =
      Workarounds.workaroundForTDFIssue89783(maxCriticalElements);
    if (ds.getSize() > maxProcessableDatasets)
    {
      WollMuxSingleton.showInfoModal(
        L.m("WollMux-Seriendruck Info"),
        L.m(
          "Ihr Seriendruckauftrag ist zu groß für das verwendete Office und wird daher in %1 Einzelaufträge aufgeteilt!",
          (int) Math.ceil((double) ds.getSize() / maxProcessableDatasets)));
    }

    // Seriendruck durchführen (ggf. aufgeteilt in mehrere "Päckchen")
    int start = 0, end = maxProcessableDatasets;
    while (!pmod.isCanceled())
    {
      if (end > ds.getSize()) end = ds.getSize();
      
      ds.setSelection(start, end);
      mergeAndShowResult(inputFile, tmpDir, type, pmod, ds, null);

      start = end;
      end += maxProcessableDatasets;
      if (start >= ds.getSize()) break;
    }

    // Aufräumen
    inputFile.delete();
    tmpDir.delete();
  }

  /**
   * Diese Methode startet den MailMergeThread für die Gesamtdokumenterzeugung aus
   * dem Hauptdokument inputFile in das tmpDir, wartet auf dessen Beendigung und
   * öffnet (falls type==OutputType.toFile) nach erfolgreichem Mergen das
   * Ergebnisdokument.
   */
  private static void mergeAndShowResult(File inputFile, File outputDir,
      OutputType type, final XPrintModel pmod, MemoryDataSource ds, String dbName)
  {
    MailMergeThread t = null;
    try
    {
      PrintModels.setStage(
        pmod,
        (ds.getSize() == ds.getSelectionSize()) ? L.m("Gesamtdokument erzeugen")
                                               : L.m("Gesamtdokumente erzeugen"));
      ProgressUpdater updater =
        new ProgressUpdater(pmod, (int) Math.ceil((double) ds.getSize()
          / countNextSets(pmod.getTextDocument())), ds.getSelectionStart());

      // Lese ausgewählten Drucker
      XPrintable xprintSD =
        (XPrintable) UnoRuntime.queryInterface(XPrintable.class,
          pmod.getTextDocument());
      String pNameSD;
      PropertyValue[] printer = null;
      if (xprintSD != null) printer = xprintSD.getPrinter();
      UnoProps printerInfo = new UnoProps(printer);
      try
      {
        pNameSD = (String) printerInfo.getPropertyValue("Name");
      }
      catch (UnknownPropertyException e)
      {
        pNameSD = "unbekannt";
      }
      t = runMailMerge(ds, outputDir, inputFile, updater, type, pNameSD);
    }
    catch (Exception e)
    {
      Logger.error(L.m("Fehler beim Starten des OOo-Seriendrucks"), e);
    }

    // Warte auf Ende des MailMerge-Threads unter Berücksichtigung von
    // pmod.isCanceled()
    while (t != null && t.isAlive())
      try
      {
        t.join(1000);
        if (pmod.isCanceled())
        {
          t.cancel();
          break;
        }
      }
      catch (InterruptedException e)
      {}
    if (pmod.isCanceled() && t.isAlive())
    {
      t.interrupt();
      Logger.debug(L.m("Der OOo-Seriendruck wurde abgebrochen"));
      // aber aufräumen tun wir noch...
    }

    if (type == OutputType.toFile && !pmod.isCanceled())
    {
      // Output-File als Template öffnen und aufräumen
      File outputFile = new File(outputDir, "output0.odt");
      boolean success = false;

      if (outputFile.exists())
      {
        try
        {
          String unoURL =
            UNO.getParsedUNOUrl(outputFile.toURI().toString()).Complete;
          Logger.debug(L.m("Öffne erzeugtes Gesamtdokument %1", unoURL));
          XComponent resultDoc = UNO.loadComponentFromURL(unoURL, true, false);

          // update LastTouchedByVersionInfo
          if (UNO.XTextDocument(resultDoc) != null)
          {
            TextDocumentModel m =
              WollMuxSingleton.getInstance().getTextDocumentModel(
                UNO.XTextDocument(resultDoc));
            m.updateLastTouchedByVersionInfo();
          }
          outputFile.delete();
          success = true;
        }
        catch (Exception e)
        {
          Logger.error(e);
        }
      }

      if (!success)
      {
        WollMuxSingleton.showInfoModal(L.m("WollMux-Seriendruck"),
          L.m("Leider konnte kein Gesamtdokument erstellt werden."));
        pmod.cancel();
      }
    }
  }

  /**
   * A optional XCancellable mail merge thread.
   * 
   * @author Jan-Marek Glogowski (ITM-I23)
   */
  private static class MailMergeThread extends Thread
  {
    private XCancellable mailMergeCancellable = null;

    private final XJob mailMerge;

    private final File outputDir;

    private final ArrayList<NamedValue> mmProps;

    MailMergeThread(XJob mailMerge, File outputDir, ArrayList<NamedValue> mmProps)
    {
      this.mailMerge = mailMerge;
      this.outputDir = outputDir;
      this.mmProps = mmProps;
    }

    public void run()
    {
      try
      {
        Logger.debug(L.m("Starting OOo-MailMerge in Verzeichnis %1", outputDir));
        // The XCancellable mail merge interface was included in LO >= 4.3.
        mailMergeCancellable =
          (XCancellable) UnoRuntime.queryInterface(XCancellable.class, mailMerge);
        if (mailMergeCancellable != null)
          Logger.debug(L.m("XCancellable interface im mailMerge-Objekt gefunden!"));
        else
          Logger.debug(L.m("KEIN XCancellable interface im mailMerge-Objekt gefunden!"));
        mailMerge.execute(mmProps.toArray(new NamedValue[mmProps.size()]));
        Logger.debug(L.m("Finished Mail Merge"));
      }
      catch (Exception e)
      {
        Logger.debug(L.m("OOo-MailMergeService fehlgeschlagen: %1", e.getMessage()));
      }
      mailMergeCancellable = null;
    }

    public synchronized void cancel()
    {
      if (mailMergeCancellable != null) mailMergeCancellable.cancel();
    }
  }

  /**
   * Übernimmt das Aktualisieren der Fortschrittsanzeige im XPrintModel pmod.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static class ProgressUpdater
  {
    private XPrintModel pmod;

    private int currentCount;

    public final int maxDatasets;

    public ProgressUpdater(XPrintModel pmod, int maxDatasets, int currentCount)
    {
      this.pmod = pmod;
      this.currentCount = currentCount;
      this.maxDatasets = maxDatasets;
      pmod.setPrintProgressMaxValue((short) maxDatasets);
      pmod.setPrintProgressValue((short) currentCount);
    }

    public void incrementProgress()
    {
      pmod.setPrintProgressValue((short) ++currentCount);
    }

    public void setMessage(String text)
    {
      this.currentCount = 0;
      pmod.setPrintMessage(text);
    }
  }

  /**
   * Implementiert einen SimulationResultsProcessor, der zugleich ein XResultSet, ein
   * XRowLocate und ein XColumnsSupplier für den OOoMailMerge ist. Dabei werden bei
   * XRowLocate als Bookmarks Werte vom Typ Integer verwendet, die den Index der
   * jeweiligen Zeile repräsentieren. Sowohl für Bookmarks, für absolute(index) als
   * auch den internen Zähler currentPos gilt: 0 repräsentiert den Datensatz vor dem
   * ersten Datensatz, 1 repräsentiert den ersten Datensatz, getSize() repräsentiert
   * den letzten existierenden Datensatz und getSize()+1 die Position "afterLast".
   * 
   * Über die Methode setSelection(start, end) kann eine Auswahl festgelegt werden,
   * die sich auf dsas Verhalten der Methoden aus XResultSet, XRowLocate,
   * XColumnsSupplier auswirkt. So kann dem MailMerge-Service eine Auswahl
   * vorgegaukelt werden (das ist notwendig, weil die im MailMerge-Service direkt per
   * Selection Property setzbare Auswahl nicht korrekt funktioniert).
   * 
   * Es gibt einen JUnit-Test zu dieser Klasse!
   */
  public static class MemoryDataSource implements XResultSet, XRowLocate,
      XColumnsSupplier, SimulationResultsProcessor
  {
    private HashMap<String, Integer> schema = null;

    private Vector<MyRow> datasets = new Vector<MyRow>();

    private MyRow emptyRow = new MyRow();

    /**
     * Wenn keine Selection gesetzt ist dann gilt: 0 ist "vor dem ersten Element", 1
     * ist das erste Element, getSize() ist das letzte Element und der Maximalwert
     * ist getSize()+1 ("nach dem letzten Element").
     * 
     * Wenn eine Selektion gesetzt ist, dann gitlt: selectionStart ist
     * "vor dem ersten Element", selectionStart+1 ist das erste Element, selectionEnd
     * ist das letzte Element und der Maximalwert ist selectionEnd+1
     * ("nach dem letzten Element").
     */
    private int currentPos = 0;

    /**
     * Entspricht ersten Element der Selektion, die Zählung beginnt mit 0.
     */
    private int selectionStart = 0;

    /**
     * Entspricht dem Element "nach dem letzen Element der Selektion".
     */
    private int selectionEnd = 0;

    /*
     * (non-Javadoc)
     * 
     * @see
     * de.muenchen.allg.itd51.wollmux.SimulationResults.SimulationResultsProcessor
     * #processSimulationResults(de.muenchen.allg.itd51.wollmux.SimulationResults)
     */
    @Override
    public void processSimulationResults(SimulationResults simRes)
    {
      if (simRes == null) return;

      HashMap<String, String> data =
        new HashMap<String, String>(simRes.getFormFieldValues());
      for (FormField field : simRes.getFormFields())
      {
        String columnName = getSpecialColumnNameForFormField(field);
        if (columnName == null) continue;
        String content = simRes.getFormFieldContent(field);

        // Checkboxen müssen über bestimmte Zeichen der Schriftart OpenSymbol
        // angenähert werden.
        if (field.getType() == FormFieldType.CheckBoxFormField)
          if (content.equalsIgnoreCase("TRUE"))
            content = "" + OPENSYMBOL_CHECKED;
          else
            content = "" + OPENSYMBOL_UNCHECKED;

        data.put(columnName, content);
      }
      try
      {
        addDataset(data);
      }
      catch (Exception e)
      {
        Logger.error(e);
      }
    }

    /**
     * Setzt die Selektion, die dem MailMerge-Service vorgegaukelt werden soll; start
     * entspricht dabei dem ersten Element (wobei die Zählung mit 0 beginnt) und end
     * dem Element "nach dem letzten Element" der Auswahl.
     */
    public void setSelection(int start, int end)
    {
      selectionStart = start;
      selectionEnd = end;
      if (selectionStart < 0) selectionStart = 0;
      if (selectionEnd > getSize()) selectionEnd = getSize();
    }

    /**
     * Liefert die Start-Position der aktuellen per setSelection(...) gesetzten
     * Auswahl zurück oder 0, wenn es keine Auswahl gibt.
     */
    public int getSelectionStart()
    {
      return selectionStart;
    }

    /**
     * Liefert die Größe der aktuell per setSelection(...) gesetzten Auswahl zurück
     * oder getSize(), wenn es keine Auswahl gibt.
     */
    public int getSelectionSize()
    {
      return selectionEnd - selectionStart;
    }

    /**
     * Fügt der Datenquelle eine neue Zeile mit den in dataset enthaltenen Daten zu.
     * Diese Daten werden nicht nur als Referenz übernommen, sondern kopiert. Die
     * Schlüssel aus dem zuerst hinzugefügten Datensatz bestimmen dabei das Schema
     * (die verfügbaren Spalten) der Datenquelle, das sich mit weiteren Datensätzen
     * nicht mehr ändert.
     */
    public void addDataset(HashMap<String, String> dataset)
    {
      if (schema == null)
      {
        schema = new HashMap<String, Integer>();
        int i = 0;
        for (Map.Entry<String, String> entry : dataset.entrySet())
        {
          schema.put(entry.getKey(), i);
          ++i;
        }
      }
      datasets.add(new MyRow(dataset, schema));
      selectionEnd = getSize();
      if (emptyRow.getElementNames().length == 0 && schema.size() > 0)
        emptyRow = new MyRow(new HashMap<String, String>(), schema);
    }

    public int getSize()
    {
      return datasets.size();
    }

    @Override
    public int compareBookmarks(Object a, Object b) throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED(a, b);
      if (a != null && b != null && a instanceof Integer && b instanceof Integer)
      {
        return ((Integer) a).compareTo((Integer) b);
      }
      throw new SQLException("Bookmark must be an Integer");
    }

    @Override
    public Object getBookmark() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return Integer.valueOf(currentPos);
    }

    @Override
    public boolean hasOrderedBookmarks() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return true;
    }

    @Override
    public int hashBookmark(Object bookmark) throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return bookmark.hashCode();
    }

    /**
     * Setzt die aktuelle Position um offset relativ zur Position von bookmark. Wird
     * auf die Positionen "beforeFirstRow" (selectionStart), bzw. "afterLastRow"
     * (selectionEnd + 1) gesetzt, liefert die Funktion false, setzt aber die
     * Position um. In allen anderen ungültigen Fällen verhindert sie einen Setzen in
     * den negativen Bereich bzw. in den Bereich über getSize()+1 und liefert false
     * zurück.
     */
    private boolean setPosRelativeToBookmark(Object bookmark, int offset)
    {
      if (bookmark != null && bookmark instanceof Integer)
      {
        int newPos = selectionStart + ((Integer) bookmark) + offset;
        if (newPos <= selectionStart)
        {
          currentPos = selectionStart;
          return false;
        }
        else if (newPos > selectionEnd)
        {
          currentPos = selectionEnd + 1;
          return false;
        }
        else
          currentPos = newPos;
        return true;
      }
      return false;
    }

    public boolean moveRelativeToBookmark(Object bookmark, int offset)
        throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED(bookmark, offset);
      return setPosRelativeToBookmark(bookmark, offset);
    }

    @Override
    public boolean moveToBookmark(Object bookmark) throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED(bookmark);
      return setPosRelativeToBookmark(bookmark, 0);
    }

    @Override
    public boolean absolute(int id) throws SQLException
    {
      if (id < 0)
        return setPosRelativeToBookmark(Integer.valueOf(getSelectionSize() + 1), id);
      else
        return setPosRelativeToBookmark(Integer.valueOf(0), id);
    }

    @Override
    public void afterLast() throws SQLException
    {
      currentPos = selectionEnd + 1;
    }

    @Override
    public void beforeFirst() throws SQLException
    {
      currentPos = selectionStart;
    }

    @Override
    public boolean first() throws SQLException
    {
      return absolute(1);
    }

    @Override
    public int getRow() throws SQLException
    {
      return currentPos - selectionStart;
    }

    @Override
    public Object getStatement() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return Any.VOID;
    }

    @Override
    public boolean isAfterLast() throws SQLException
    {
      return currentPos == selectionEnd + 1;
    }

    @Override
    public boolean isBeforeFirst() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return currentPos == selectionStart;
    }

    @Override
    public boolean isFirst() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return currentPos == selectionStart + 1;
    }

    @Override
    public boolean isLast() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return currentPos == selectionEnd;
    }

    @Override
    public boolean last() throws SQLException
    {
      return absolute(-1);
    }

    @Override
    public boolean next() throws SQLException
    {
      return setPosRelativeToBookmark(Integer.valueOf(currentPos - selectionStart),
        1);
    }

    @Override
    public boolean previous() throws SQLException
    {
      return setPosRelativeToBookmark(Integer.valueOf(currentPos - selectionStart),
        -1);
    }

    @Override
    public void refreshRow() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
    }

    @Override
    public boolean relative(int offset) throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED(offset);
      return setPosRelativeToBookmark(Integer.valueOf(currentPos - selectionStart),
        offset);
    }

    @Override
    public boolean rowDeleted() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return false;
    }

    @Override
    public boolean rowInserted() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return false;
    }

    @Override
    public boolean rowUpdated() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return false;
    }

    @Override
    public XNameAccess getColumns()
    {
      if (currentPos > selectionStart && currentPos <= selectionEnd)
      {
        return datasets.get(currentPos - 1);
      }
      return emptyRow;
    }
  }

  public static class MyRow implements XNameAccess
  {
    private String[] data;

    private HashMap<String, Integer> schema;

    /**
     * Erzeugt eine leere Row ohne Spalten.
     */
    MyRow()
    {
      this.data = new String[] {};
      this.schema = new HashMap<String, Integer>();
    }
    
    MyRow(HashMap<String, String> dataset, HashMap<String, Integer> schema)
    {
      this.schema = schema;
      this.data = new String[schema.size()];
      for (Map.Entry<String, Integer> entry : schema.entrySet())
      {
        data[entry.getValue()] = dataset.get(entry.getKey());
      }
    }

    @Override
    public boolean hasElements()
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return schema.size() > 0;
    }

    @Override
    public Type getElementType()
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return Type.ANY;
    }

    @Override
    public boolean hasByName(String columnName)
    {
      return schema.containsKey(columnName);
    }

    @Override
    public String[] getElementNames()
    {
      String[] names = new String[0];
      names = new String[schema.size()];
      schema.keySet().toArray(names);
      return names;
    }

    @Override
    public Object getByName(String columnName) throws NoSuchElementException,
        WrappedTargetException
    {
      Integer index = schema.get(columnName);
      if (index == null) throw new NoSuchElementException(columnName);
      String value = data[index];
      if (value == null) value = "";
      return new MyColumn(value);
    }
  }

  public static class MyColumn implements XColumn, XPropertySet
  {
    private String value;

    public MyColumn(String value)
    {
      this.value = value;
    }

    @Override
    public Object getPropertyValue(String name) throws UnknownPropertyException,
        WrappedTargetException
    {
      if ("Type".equalsIgnoreCase(name))
        return DataType.VARCHAR;
      else
      {
        NOT_YET_IMPLEMENTED_OR_REQUIRED(name);
        throw new UnknownPropertyException(name);
      }
    }

    @Override
    public String getString() throws SQLException
    {
      return value;
    }

    // *****************************************************************************
    // Jetzt kommen alle nicht implementierten Methoden. Wir müssen sie auch nicht
    // implementieren, da sie durch obige Rahmenbedingungen (DataType.VARCHAR) nicht
    // aufgerufen werden.

    @Override
    public void setPropertyValue(String arg0, Object arg1)
        throws UnknownPropertyException, PropertyVetoException,
        IllegalArgumentException, WrappedTargetException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED(arg0, arg1);
    }

    @Override
    public void removeVetoableChangeListener(String arg0,
        XVetoableChangeListener arg1) throws UnknownPropertyException,
        WrappedTargetException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED(arg0, arg1);
    }

    @Override
    public void removePropertyChangeListener(String arg0,
        XPropertyChangeListener arg1) throws UnknownPropertyException,
        WrappedTargetException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED(arg0, arg1);
    }

    @Override
    public XPropertySetInfo getPropertySetInfo()
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public void addVetoableChangeListener(String arg0, XVetoableChangeListener arg1)
        throws UnknownPropertyException, WrappedTargetException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED(arg0, arg1);
    }

    @Override
    public void addPropertyChangeListener(String arg0, XPropertyChangeListener arg1)
        throws UnknownPropertyException, WrappedTargetException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED(arg0, arg1);
    }

    @Override
    public XArray getArray() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public XInputStream getBinaryStream() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public XBlob getBlob() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public boolean getBoolean() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return false;
    }

    @Override
    public byte getByte() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return 0;
    }

    @Override
    public byte[] getBytes() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public XInputStream getCharacterStream() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public XClob getClob() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public Date getDate() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public double getDouble() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return 0;
    }

    @Override
    public float getFloat() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return 0;
    }

    @Override
    public int getInt() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return 0;
    }

    @Override
    public long getLong() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return 0;
    }

    @Override
    public Object getObject(XNameAccess arg0) throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED(arg0);
      return null;
    }

    @Override
    public XRef getRef() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public short getShort() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return 0;
    }

    @Override
    public Time getTime() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public DateTime getTimestamp() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return null;
    }

    @Override
    public boolean wasNull() throws SQLException
    {
      NOT_YET_IMPLEMENTED_OR_REQUIRED();
      return false;
    }
  };

  private static void NOT_YET_IMPLEMENTED_OR_REQUIRED(Object... args)
  {
    StackTraceElement[] st = Thread.currentThread().getStackTrace();
    StringBuffer buf = new StringBuffer();
    if (st.length >= 3)
    {
      buf.append(st[3]);
      for (Object arg : args)
      {
        if (buf.length() == 0)
          buf.append(", args=");
        else
          buf.append(", ");
        buf.append("'" + arg + "'");
      }
    }
    Logger.error(L.m("Aufruf einer nicht implementierten Funktion: %1",
      buf.toString()));
  }

  /**
   * Im Rahmen von trac#14905 wurde in den WollMux ein Workaround zur Umgehung der
   * 16Bit-Limitierungen in OOo/AOO/LO eingebaut. Dabei werden manche Container in
   * OOo/AOO/LO noch per 16Bit-adressiert, womit es dann nur 65536 Elemente in diesen
   * Containern geben kann. Konkret verhält sich LO 4.1.6 so, dass es einfriert, wenn
   * z.B. ein Hauptdokument 32 Rahmen enthält und 2048 mal in ein Gesamtdokument
   * geschrieben werden soll. Um dies zu verhindern, erfasst der WollMux in dieser
   * Klasse vor dem Druckauftrag eine Statistik über das Hauptdokument. So kann durch
   * Auswertung dieser Statisktik ein Einfrieren des Seriendrucks umgangen werden,
   * bzw. der Druckauftrag in kleinere Druckaufträge aufgeteilt werden.
   * 
   * @author Christoph Lutz (CIB software GmbH)
   */
  public static class DocStatistics
  {
    private int containedTextframes = 0;

    private int containedSections = 0;

    private int containedPageStyles = 0;

    private int containedTables = 0;

    public int getContainedTextframes()
    {
      return containedTextframes;
    }

    public int getContainedSections()
    {
      return containedSections;
    }

    public int getContainedPageStyles()
    {
      return containedPageStyles;
    }

    public int getContainedTables()
    {
      return containedTables;
    }

    public void countElements(XComponent doc)
    {
      // count Sections:
      XTextSectionsSupplier tss = UNO.XTextSectionsSupplier(doc);
      if (tss != null)
      {
        String[] names = tss.getTextSections().getElementNames();
        containedSections = names.length;
      }

      // count TextFrames:
      XTextFramesSupplier tfs = UNO.XTextFramesSupplier(doc);
      if (tfs != null)
      {
        String[] names = tfs.getTextFrames().getElementNames();
        containedTextframes = names.length;
      }

      // count TextTables
      XTextTablesSupplier tts =
        UnoRuntime.queryInterface(XTextTablesSupplier.class, doc);
      if (tts != null)
      {
        String[] names = tts.getTextTables().getElementNames();
        containedTables = names.length;
      }

      // count (really used) PageStyles
      try
      {
        XTextDocument tdoc = UNO.XTextDocument(doc);
        XParagraphCursor cursor =
          UNO.XParagraphCursor(tdoc.getText().createTextCursorByRange(
            tdoc.getText().getStart()));
        HashSet<String> usedPageStyles = new HashSet<String>();
        do
        {
          Object pageStyleName = UNO.getProperty(cursor, "PageStyleName");
          if (pageStyleName != null) usedPageStyles.add(pageStyleName.toString());
        } while (cursor.gotoNextParagraph(false));

        containedPageStyles = usedPageStyles.size();
      }
      catch (java.lang.Exception e)
      {}
    }

    public String toString()
    {
      return this.getClass().getName() + "(Textframes=" + containedTextframes
        + ", Sections=" + containedSections + ", PageStyles=" + containedPageStyles
        + ", TextTables=" + containedTables + ")";
    }
  }

  /**
   * Erzeugt das aus origDoc abgeleitete, für den OOo-Seriendruck heranzuziehende
   * Input-Dokument im Verzeichnis tmpDir und nimmt alle notwendigen Anpassungen vor,
   * damit der Seriendruck über die temporäre Datenbank dbName korrekt und möglichst
   * performant funktioniert, und liefert dieses zurück.
   * 
   * @author Christoph Lutz (D-III-ITD-D101) TESTED
   */
  private static File createAndAdjustInputFile(File tmpDir, XTextDocument origDoc,
      DocStatistics s)
  {
    // Aktuelles Dokument speichern als neues input-Dokument
    if (origDoc == null) return null;
    File inputFile = new File(tmpDir, "input.odt");
    String url = UNO.getParsedUNOUrl(inputFile.toURI().toString()).Complete;
    XStorable xStorable = UNO.XStorable(origDoc);
    if (xStorable != null)
    {
      try
      {
        xStorable.storeToURL(url, new PropertyValue[] {});
      }
      catch (IOException e)
      {
        Logger.error(
          L.m("Kann temporäres Eingabedokument für den OOo-Seriendruck nicht erzeugen"),
          e);
        return null;
      }
    }
    else
    {
      return null;
    }

    // Neues input-Dokument öffnen. Achtung: Normalerweise würde der
    // loadComponentFromURL den WollMux veranlassen, das Dokument zu interpretieren
    // (und damit zu verarbeiten). Da das bei diesem temporären Dokument nicht
    // erwünscht ist, erkennt der WollMux in
    // d.m.a.i.wollmux.event.GlobalEventListener.isTempMailMergeDocument(XModel
    // compo) über den Pfad der Datei dass es sich um ein temporäres Dokument handelt
    // und dieses nicht bearbeitet werden soll.
    XComponent tmpDoc = null;
    try
    {
      tmpDoc = UNO.loadComponentFromURL(url, false, false, true);
    }
    catch (Exception e)
    {
      return null;
    }

    // neues input-Dokument bearbeiten/anpassen
    addDatabaseFieldsForInsertFormValueBookmarks(UNO.XTextDocument(tmpDoc),
      DATA_SOURCE_NAME);
    adjustDatabaseAndInputUserFields(tmpDoc, DATA_SOURCE_NAME);
    removeAllBookmarks(tmpDoc);
    removeHiddenSections(tmpDoc);
    SachleitendeVerfuegung.deMuxSLVStyles(UNO.XTextDocument(tmpDoc));
    removeWollMuxMetadata(UNO.XTextDocument(tmpDoc));

    // Dokumentstatistik erheben (wenn s != null)
    if (s != null)
    {
      s.countElements(tmpDoc);
    }

    // neues input-Dokument speichern und schließen
    if (UNO.XStorable(tmpDoc) != null)
    {
      try
      {
        UNO.XStorable(tmpDoc).store();
      }
      catch (IOException e)
      {
        inputFile = null;
      }
    }
    else
    {
      inputFile = null;
    }

    boolean closed = false;
    if (UNO.XCloseable(tmpDoc) != null) do
    {
      try
      {
        UNO.XCloseable(tmpDoc).close(true);
        closed = true;
      }
      catch (CloseVetoException e)
      {
        try
        {
          Thread.sleep(2000);
        }
        catch (InterruptedException e1)
        {}
      }
    } while (closed == false);

    return inputFile;
  }

  /**
   * Entfernt alle Metadaten des WollMux aus dem Dokument doc die nicht reine
   * Infodaten des WollMux sind (wie z.B. WollMuxVersion, OOoVersion) um
   * sicherzustellen, dass der WollMux das Gesamtdokument nicht interpretiert.
   * 
   * @author Christoph Lutz (D-III-ITD-D101) TESTED
   */
  private static void removeWollMuxMetadata(XTextDocument doc)
  {
    if (doc == null) return;
    PersistentDataContainer c = PersistentData.createPersistentDataContainer(doc);
    for (DataID dataId : DataID.values())
      c.removeData(dataId);
    c.flush();
  }

  /**
   * Hebt alle unsichtbaren TextSections (Bereiche) in Dokument tmpDoc auf, wobei bei
   * auch der Inhalt entfernt wird. Das Entfernen der unsichtbaren Bereiche dient zur
   * Verbesserung der Performance, das Löschen der Bereichsinhalte ist notwendig,
   * damit das erzeugte Gesamtdokument korrekt dargestellt wird (hier habe ich wilde
   * Textverschiebungen beobachtet, die so vermieden werden sollen).
   * 
   * Bereiche sind auch ein möglicher Auslöser von allen möglichen falsch gesetzten
   * Seitenumbrüchen (siehe z.B. Issue:
   * http://openoffice.org/bugzilla/show_bug.cgi?id=73229)
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static void removeHiddenSections(XComponent tmpDoc)
  {
    XTextSectionsSupplier tss = UNO.XTextSectionsSupplier(tmpDoc);
    if (tss == null) return;

    for (String name : tss.getTextSections().getElementNames())
    {
      try
      {
        XTextSection section =
          UNO.XTextSection(tss.getTextSections().getByName(name));
        if (Boolean.FALSE.equals(UNO.getProperty(section, "IsVisible")))
        {
          // Inhalt der Section löschen und Section aufheben:
          section.getAnchor().setString("");
          section.getAnchor().getText().removeTextContent(section);
        }
      }
      catch (Exception e)
      {
        Logger.error(e);
      }
    }
  }

  /**
   * Aufgrund eines Bugs in OOo führen Bookmarks zu einer Verlangsamung des
   * Seriendruck in der Komplexität O(n^2) und werden hier in dieser Methode alle aus
   * dem Dokument tmpDoc gelöscht. Bookmarks sollten im Ergebnisdokument sowieso
   * nicht mehr benötigt werden und sind damit aus meiner Sicht überflüssig.
   * 
   * Sollte irgendjemand irgendwann zu der Meinung kommen, dass die Bookmarks im
   * Dokument bleiben müssen, so müssen zumindest die Bookmarks von
   * WollMux-Dokumentkommandos gelöscht werden, damit sie nicht noch einmal durch den
   * WollMux bearbeitet werden.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static void removeAllBookmarks(XComponent tmpDoc)
  {
    if (UNO.XBookmarksSupplier(tmpDoc) != null)
    {
      XNameAccess xna = UNO.XBookmarksSupplier(tmpDoc).getBookmarks();
      for (String name : xna.getElementNames())
      {
        XTextContent bookmark = null;
        try
        {
          bookmark = UNO.XTextContent(xna.getByName(name));
        }
        catch (NoSuchElementException e)
        {
          continue;
        }
        catch (Exception e)
        {
          Logger.error(e);
        }

        if (bookmark != null) try
        {
          bookmark.getAnchor().getText().removeTextContent(bookmark);
        }
        catch (NoSuchElementException e1)
        {
          Logger.error(e1);
        }
      }
    }
  }

  /**
   * Fügt dem Dokument doc für alle enthaltenen insertFormValue-Bookmarks zugehörige
   * OOo-Seriendruckfelder mit Verweis auf die Datenbank dbName hinzu.
   * 
   * @author Christoph Lutz (D-III-ITD-D101) TESTED
   */
  private static void addDatabaseFieldsForInsertFormValueBookmarks(
      XTextDocument doc, String dbName)
  {
    DocumentCommands cmds = new DocumentCommands(UNO.XBookmarksSupplier(doc));
    cmds.update();
    HashMap<String, FormField> bookmarkNameToFormField =
      new HashMap<String, FormField>();
    for (DocumentCommand cmd : cmds)
    {
      if (cmd instanceof InsertFormValue)
      {
        InsertFormValue ifvCmd = (InsertFormValue) cmd;
        FormField field =
          FormFieldFactory.createFormField(doc, ifvCmd, bookmarkNameToFormField);
        if (field == null) continue;
        field.setCommand(ifvCmd);

        String columnName = getSpecialColumnNameForFormField(field);
        if (columnName == null) columnName = ifvCmd.getID();
        try
        {
          XDependentTextField dbField =
            createDatabaseField(UNO.XMultiServiceFactory(doc), dbName, TABLE_NAME,
              columnName);
          if (dbField == null) continue;

          ifvCmd.insertTextContentIntoBookmark(dbField, true);

          // Checkboxen müssen über bestimmte Zeichen der Schriftart OpenSymbol
          // angenähert werden.
          if (field.getType() == FormFieldType.CheckBoxFormField)
            UNO.setProperty(ifvCmd.getTextCursor(), "CharFontName", "OpenSymbol");
        }
        catch (Exception e)
        {
          Logger.error(e);
        }
      }
    }
  }

  /**
   * Liefert zum Formularfeld field unter Berücksichtigung des Feld-Typs und evtl.
   * gesetzter Trafos eine eindeutige Bezeichnung für die Datenbankspalte in die der
   * Wert des Formularfeldes geschrieben ist bzw. aus der der Wert des Formularfeldes
   * wieder ausgelesen werden kann oder null, wenn das Formularfeld über einen
   * primitiven Spaltennamen (der nur aus einer in den setValues gesetzten IDs
   * besteht) gefüllt werden kann.
   * 
   * @author Christoph Lutz (D-III-ITD-D101) TESTED
   */
  private static String getSpecialColumnNameForFormField(FormField field)
  {
    String trafo = field.getTrafoName();
    String id = field.getId();

    if (field.getType() == FormFieldType.CheckBoxFormField && id != null
      && trafo != null)
      return COLUMN_PREFIX_CHECKBOX_FUNCTION + SEP + id + SEP + trafo;

    else if (field.getType() == FormFieldType.CheckBoxFormField && id != null
      && trafo == null)
      return COLUMN_PREFIX_CHECKBOX_FUNCTION + SEP + id;

    else if (field.singleParameterTrafo() && id != null && trafo != null)
      return COLUMN_PREFIX_SINGLE_PARAMETER_FUNCTION + SEP + id + SEP + trafo;

    else if (!field.singleParameterTrafo() && trafo != null)
      return COLUMN_PREFIX_MULTI_PARAMETER_FUNCTION + SEP + trafo;

    return null;
  }

  /**
   * Passt bereits enthaltene OOo-Seriendruckfelder und Nächster-Datensatz-Felder in
   * tmpDoc so an, dass sie über die Datenbank dbName befüllt werden und ersetzt
   * InputUser-Felder durch entsprechende OOo-Seriendruckfelder.
   * 
   * @author Christoph Lutz (D-III-ITD-D101) TESTED
   */
  private static void adjustDatabaseAndInputUserFields(XComponent tmpDoc,
      String dbName)
  {
    if (UNO.XTextFieldsSupplier(tmpDoc) != null)
    {
      XEnumeration xenum =
        UNO.XTextFieldsSupplier(tmpDoc).getTextFields().createEnumeration();
      while (xenum.hasMoreElements())
      {
        XDependentTextField tf = null;
        try
        {
          tf = UNO.XDependentTextField(xenum.nextElement());
        }
        catch (Exception e)
        {
          continue;
        }

        // Database-Felder anpassen auf temporäre Datenquelle/Tabelle
        if (UNO.supportsService(tf, "com.sun.star.text.TextField.Database"))
        {
          XPropertySet master = tf.getTextFieldMaster();
          UNO.setProperty(master, "DataBaseName", dbName);
          UNO.setProperty(master, "DataTableName", TABLE_NAME);
        }

        // "Nächster Datensatz"-Felder anpassen auf temporäre Datenquelle/Tabelle
        if (UNO.supportsService(tf, "com.sun.star.text.TextField.DatabaseNextSet"))
        {
          UNO.setProperty(tf, "DataBaseName", dbName);
          UNO.setProperty(tf, "DataTableName", TABLE_NAME);
        }

        // InputUser-Felder ersetzen durch entsprechende Database-Felder
        else if (UNO.supportsService(tf, "com.sun.star.text.TextField.InputUser"))
        {
          String content = "";
          try
          {
            content = AnyConverter.toString(UNO.getProperty(tf, "Content"));
          }
          catch (IllegalArgumentException e)
          {}

          String trafo = TextDocumentModel.getFunctionNameForUserFieldName(content);
          if (trafo != null)
          {
            try
            {
              XDependentTextField dbField =
                createDatabaseField(UNO.XMultiServiceFactory(tmpDoc), dbName,
                  TABLE_NAME, COLUMN_PREFIX_MULTI_PARAMETER_FUNCTION + SEP + trafo);
              tf.getAnchor().getText().insertTextContent(tf.getAnchor(), dbField,
                true);
            }
            catch (Exception e)
            {
              Logger.error(e);
            }
          }
        }
      }
    }
  }

  /**
   * Zählt die Anzahl an "Nächster Datensatz"-Feldern zur Berechnung der Gesamtzahl
   * der zu verarbeitenden Dokumente.
   * 
   * @author Ignaz Forster (ITM-I23)
   */
  private static int countNextSets(XComponent doc)
  {
    int numberOfNextSets = 1;
    if (UNO.XTextFieldsSupplier(doc) != null)
    {
      XEnumeration xenum =
        UNO.XTextFieldsSupplier(doc).getTextFields().createEnumeration();
      while (xenum.hasMoreElements())
      {
        XDependentTextField tf = null;
        try
        {
          tf = UNO.XDependentTextField(xenum.nextElement());
        }
        catch (Exception e)
        {
          continue;
        }

        if (UNO.supportsService(tf, "com.sun.star.text.TextField.DatabaseNextSet"))
        {
          numberOfNextSets++;
        }
      }
    }
    return numberOfNextSets;
  }

  /**
   * Erzeugt über die Factory factory ein neues OOo-Seriendruckfeld, das auf die
   * Datenbank dbName, die Tabelle tableName und die Spalte columnName verweist und
   * liefert dieses zurück.
   * 
   * @throws Exception
   *           Wenn die Factory das Feld nicht erzeugen kann.
   * @throws IllegalArgumentException
   *           Wenn irgendetwas mit den Attributen dbName, tableName oder columnName
   *           nicht stimmt.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static XDependentTextField createDatabaseField(
      XMultiServiceFactory factory, String dbName, String tableName,
      String columnName) throws Exception, IllegalArgumentException
  {
    XDependentTextField dbField =
      UNO.XDependentTextField(factory.createInstance("com.sun.star.text.TextField.Database"));
    XPropertySet m =
      UNO.XPropertySet(factory.createInstance("com.sun.star.text.FieldMaster.Database"));
    UNO.setProperty(m, "DataBaseName", dbName);
    UNO.setProperty(m, "DataTableName", tableName);
    UNO.setProperty(m, "DataColumnName", columnName);
    dbField.attachTextFieldMaster(m);
    return dbField;
  }

  /**
   * Deregistriert die Datenbank dbName aus der Liste der Datenbanken (wie z.B. über
   * Extras->Optionen->Base/Datenbanken einsehbar) und löscht das zugehörige in
   * tmpDir enthaltene .odb-File von der Platte.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static void removeTempDatasource(String dbName, File tmpDir)
  {
    XSingleServiceFactory dbContext =
      UNO.XSingleServiceFactory(UNO.createUNOService("com.sun.star.sdb.DatabaseContext"));
    XNamingService naming = UNO.XNamingService(dbContext);
    if (naming != null) try
    {
      naming.revokeObject(dbName);
    }
    catch (Exception e)
    {
      Logger.error(e);
    }
    new File(tmpDir, DATASOURCE_ODB_FILENAME).delete();
  }

  /**
   * Steuert den Ausgabetyp beim OOo-Seriendruck.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static enum OutputType {
    toFile,
    toPrinter;
  }

  /**
   * Startet die Ausführung des Seriendrucks in ein Gesamtdokument mit dem
   * c.s.s.text.MailMergeService in einem eigenen Thread und liefert diesen zurück.
   * 
   * @param dbName
   *          Name der Datenbank, die für den Seriendruck verwendet werden soll.
   * @param outputDir
   *          Directory in dem das Ergebnisdokument abgelegt werden soll.
   * @param inputFile
   *          Hauptdokument, das für den Seriendruck herangezogen wird.
   * @param progress
   *          Ein ProgressUpdater, der über den Bearbeitungsfortschritt informiert
   *          wird.
   * @param printerName
   *          Drucker fuer den Seriendruck
   * @throws Exception
   *           falls der MailMergeService nicht erzeugt werden kann.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static MailMergeThread runMailMerge(final MemoryDataSource ds,
      final File outputDir, File inputFile, final ProgressUpdater progress,
      final OutputType type, String printerName) throws Exception
  {
    final XJob mailMerge =
      (XJob) UnoRuntime.queryInterface(XJob.class,
        UNO.xMCF.createInstanceWithContext("com.sun.star.text.MailMerge",
          UNO.defaultContext));

    // Register MailMergeEventListener
    XMailMergeBroadcaster xmmb =
      (XMailMergeBroadcaster) UnoRuntime.queryInterface(XMailMergeBroadcaster.class,
        mailMerge);
    xmmb.addMailMergeEventListener(new XMailMergeListener()
    {
      int count = 0;

      final long start = System.currentTimeMillis();

      public void notifyMailMergeEvent(MailMergeEvent arg0)
      {
        if (progress != null) progress.incrementProgress();
        count++;
        Logger.debug2(L.m("OOo-MailMerger: verarbeite Datensatz %1 (%2 ms)", count,
          (System.currentTimeMillis() - start)));
        if (count >= progress.maxDatasets && type == OutputType.toPrinter)
        {
          progress.setMessage(L.m("Sende Druckauftrag - bitte warten..."));
        }
      }
    });

    final ArrayList<NamedValue> mmProps = new ArrayList<NamedValue>();
    mmProps.add(new NamedValue("ResultSet", ds));
    mmProps.add(new NamedValue("DataSourceName", DATA_SOURCE_NAME));
    mmProps.add(new NamedValue("CommandType", CommandType.TABLE));
    mmProps.add(new NamedValue("Command", TABLE_NAME));
    mmProps.add(new NamedValue("DocumentURL",
      UNO.getParsedUNOUrl(inputFile.toURI().toString()).Complete));
    mmProps.add(new NamedValue("OutputURL",
      UNO.getParsedUNOUrl(outputDir.toURI().toString()).Complete));
    if (type == OutputType.toFile)
    {
      mmProps.add(new NamedValue("SaveAsSingleFile", Boolean.TRUE));
      mmProps.add(new NamedValue("OutputType", MailMergeType.FILE));
      mmProps.add(new NamedValue("FileNameFromColumn", Boolean.FALSE));
      mmProps.add(new NamedValue("FileNamePrefix", "output"));
    }
    else if (type == OutputType.toPrinter)
    {
      mmProps.add(new NamedValue("OutputType", MailMergeType.PRINTER));
      mmProps.add(new NamedValue("SinglePrintJobs", Boolean.FALSE));
      // jgm,07.2013: setze ausgewaehlten Drucker
      if (printerName != null && printerName.length() > 0)
      {
        PropertyValue[] printOpts = new PropertyValue[1];
        printOpts[0] = new PropertyValue();
        printOpts[0].Name = "PrinterName";
        printOpts[0].Value = printerName;
        Logger.debug(L.m("Seriendruck - Setze Drucker: %1", printerName));
        mmProps.add(new NamedValue("PrintOptions", printOpts));
      }
      // jgm ende
    }
    MailMergeThread t = new MailMergeThread(mailMerge, outputDir, mmProps);
    t.start();
    return t;
  }

  /**
   * Erzeugt ein neues temporäres Directory mit dem Aufbau
   * "<TEMP_WOLLMUX_MAILMERGE_PREFIX>xxx" (wobei xxx eine garantiert 3-stellige Zahl
   * ist), in dem sämtliche (temporäre) Dateien für den Seriendruck abgelegt werden
   * und liefert dieses zurück.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static File createMailMergeTempdir()
  {
    File sysTmp = new File(System.getProperty("java.io.tmpdir"));
    File tmpDir;
    do
    {
      // +100 um eine 3-stellige Zufallszahl zu garantieren
      tmpDir =
        new File(sysTmp, TEMP_WOLLMUX_MAILMERGE_PREFIX
          + (new Random().nextInt(899) + 100));
    } while (!tmpDir.mkdir());
    return tmpDir;
  }

  /**
   * Testmethode
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static void main(String[] args)
  {
    String pNameSD = "HP1010"; // Drucker Name
    try
    {
      UNO.init();

      File tmpDir = createMailMergeTempdir();

      MemoryDataSource ds = new MemoryDataSource();

      File inputFile =
        createAndAdjustInputFile(tmpDir,
          UNO.XTextDocument(UNO.desktop.getCurrentComponent()), null);

      runMailMerge(ds, tmpDir, inputFile, null, OutputType.toFile, pNameSD);

      removeTempDatasource(null, tmpDir);

      inputFile.delete();

      // Output-File als Template öffnen und aufräumen
      File outputFile = new File(tmpDir, "output0.odt");
      UNO.loadComponentFromURL(
        UNO.getParsedUNOUrl(outputFile.toURI().toString()).Complete, true, false);
      outputFile.delete();
      tmpDir.delete();

    }
    catch (Exception e)
    {
      e.printStackTrace();
      System.exit(1);
    }
    System.exit(0);
  }
}
