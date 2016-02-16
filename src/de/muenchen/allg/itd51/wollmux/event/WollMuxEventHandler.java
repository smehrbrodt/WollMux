/*
 * Dateiname: WollMuxEventHandler.java
 * Projekt  : WollMux
 * Funktion : Ermöglicht die Einstellung neuer WollMuxEvents in die EventQueue.
 * 
 * Copyright (c) 2010-2015 Landeshauptstadt München
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
 * 24.10.2005 | LUT | Erstellung als EventHandler.java
 * 01.12.2005 | BNK | +on_unload() das die Toolbar neu erzeugt (böser Hack zum 
 *                  | Beheben des Seitenansicht-Toolbar-Verschwindibus-Problems)
 *                  | Ausgabe des hashCode()s in den Debug-Meldungen, um Events 
 *                  | Objekten zuordnen zu können beim Lesen des Logfiles
 * 27.03.2005 | LUT | neues Kommando openDocument
 * 21.04.2006 | LUT | +ConfigurationErrorException statt NodeNotFoundException bei
 *                    fehlendem URL-Attribut in Textfragmenten
 * 06.06.2006 | LUT | + Ablösung der Event-Klasse durch saubere Objektstruktur
 *                    + Überarbeitung vieler Fehlermeldungen
 *                    + Zeilenumbrüche in showInfoModal, damit keine unlesbaren
 *                      Fehlermeldungen mehr ausgegeben werden.
 * 16.12.2009 | ERT | Cast XTextContent-Interface entfernt
 * 03.03.2010 | ERT | getBookmarkNamesStartingWith nach UnoHelper TextDocument verschoben
 * 03.03.2010 | ERT | Verhindern von Überlagern von insertFrag-Bookmarks
 * 23.03.2010 | ERT | [R59480] Meldung beim Ausführen von "Textbausteinverweis einfügen"
 * 02.06.2010 | BED | +handleSaveTempAndOpenExt
 * 08.05.2012 | jub | fakeSymLink behandlung eingebaut: auflösung und test der FRAG_IDs berücksichtigt
 *                    auch die möglichkeit, dass im conifg file auf einen fake SymLink verwiesen wird.
 * 11.12.2012 | jub | fakeSymLinks werden doch nicht gebraucht; wieder aus dem code entfernt                   
 *
 * -------------------------------------------------------------------
 *
 * @author Christoph Lutz (D-III-ITD 5.1)
 * 
 */
package de.muenchen.allg.itd51.wollmux.event;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Frame;
import java.awt.GridLayout;
import java.awt.HeadlessException;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.io.StringReader;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Vector;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.AbstractAction;
import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.filechooser.FileFilter;

import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyValue;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.document.XEventListener;
import com.sun.star.form.binding.InvalidBindingStateException;
import com.sun.star.frame.XDispatch;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XFrames;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.RuntimeException;
import com.sun.star.util.XStringSubstitution;
import com.sun.star.view.DocumentZoomType;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoProps;
import de.muenchen.allg.afid.UnoService;
import de.muenchen.allg.itd51.parser.ConfigThingy;
import de.muenchen.allg.itd51.parser.NodeNotFoundException;
import de.muenchen.allg.itd51.wollmux.Bookmark;
import de.muenchen.allg.itd51.wollmux.ConfigurationErrorException;
import de.muenchen.allg.itd51.wollmux.DocumentCommand;
import de.muenchen.allg.itd51.wollmux.DocumentCommandInterpreter;
import de.muenchen.allg.itd51.wollmux.DocumentCommands;
import de.muenchen.allg.itd51.wollmux.DocumentManager;
import de.muenchen.allg.itd51.wollmux.DocumentManager.TextDocumentInfo;
import de.muenchen.allg.itd51.wollmux.GlobalFunctions;
import de.muenchen.allg.itd51.wollmux.L;
import de.muenchen.allg.itd51.wollmux.Logger;
import de.muenchen.allg.itd51.wollmux.ModalDialogs;
import de.muenchen.allg.itd51.wollmux.OpenExt;
import de.muenchen.allg.itd51.wollmux.PersoenlicheAbsenderliste;
import de.muenchen.allg.itd51.wollmux.SachleitendeVerfuegung;
import de.muenchen.allg.itd51.wollmux.TextDocumentModel;
import de.muenchen.allg.itd51.wollmux.TextModule;
import de.muenchen.allg.itd51.wollmux.TimeoutException;
import de.muenchen.allg.itd51.wollmux.VisibilityElement;
import de.muenchen.allg.itd51.wollmux.VisibleTextFragmentList;
import de.muenchen.allg.itd51.wollmux.WMCommandsFailedException;
import de.muenchen.allg.itd51.wollmux.WollMuxFehlerException;
import de.muenchen.allg.itd51.wollmux.WollMuxFiles;
import de.muenchen.allg.itd51.wollmux.WollMuxSingleton;
import de.muenchen.allg.itd51.wollmux.WollMuxSingleton.InvalidIdentifierException;
import de.muenchen.allg.itd51.wollmux.Workarounds;
import de.muenchen.allg.itd51.wollmux.XPALChangeEventListener;
import de.muenchen.allg.itd51.wollmux.XPrintModel;
import de.muenchen.allg.itd51.wollmux.db.DJDataset;
import de.muenchen.allg.itd51.wollmux.db.DJDatasetListElement;
import de.muenchen.allg.itd51.wollmux.db.Dataset;
import de.muenchen.allg.itd51.wollmux.db.DatasourceJoiner;
import de.muenchen.allg.itd51.wollmux.db.QueryResults;
import de.muenchen.allg.itd51.wollmux.dialog.AbsenderAuswaehlen;
import de.muenchen.allg.itd51.wollmux.dialog.Common;
import de.muenchen.allg.itd51.wollmux.dialog.Dialog;
import de.muenchen.allg.itd51.wollmux.dialog.PersoenlicheAbsenderlisteVerwalten;
import de.muenchen.allg.itd51.wollmux.dialog.formmodel.FormModel;
import de.muenchen.allg.itd51.wollmux.dialog.formmodel.InvalidFormDescriptorException;
import de.muenchen.allg.itd51.wollmux.dialog.formmodel.MultiDocumentFormModel;
import de.muenchen.allg.itd51.wollmux.dialog.formmodel.SingleDocumentFormModel;
import de.muenchen.allg.itd51.wollmux.dialog.mailmerge.MailMergeNew;
import de.muenchen.allg.itd51.wollmux.former.FormularMax4kController;
import de.muenchen.allg.itd51.wollmux.func.Function;
import de.muenchen.allg.itd51.wollmux.func.FunctionFactory;
import de.muenchen.allg.itd51.wollmux.func.FunctionLibrary;
import de.muenchen.allg.itd51.wollmux.func.Values;
import de.muenchen.allg.ooo.TextDocument;

/**
 * Ermöglicht die Einstellung neuer WollMuxEvents in die EventQueue.
 * 
 * @author Christoph Lutz (D-III-ITD 5.1)
 */
public class WollMuxEventHandler
{
  /**
   * Name des OnWollMuxProcessingFinished-Events.
   */
  public static final String ON_WOLLMUX_PROCESSING_FINISHED =
    "OnWollMuxProcessingFinished";

  /**
   * Mit dieser Methode ist es möglich die Entgegennahme von Events zu blockieren.
   * Alle eingehenden Events werden ignoriert, wenn accept auf false gesetzt ist und
   * entgegengenommen, wenn accept auf true gesetzt ist.
   * 
   * @param accept
   */
  public static void setAcceptEvents(boolean accept)
  {
    EventProcessor.getInstance().setAcceptEvents(accept);
  }

  /**
   * Der EventProcessor sorgt für eine synchronisierte Verarbeitung aller
   * Wollmux-Events. Alle Events werden in eine synchronisierte eventQueue
   * hineingepackt und von einem einzigen eventProcessingThread sequentiell
   * abgearbeitet.
   * 
   * @author lut
   */
  public static class EventProcessor
  {
    /**
     * Gibt an, ob der EventProcessor überhaupt events entgegennimmt. Ist
     * acceptEvents=false, werden alle Events ignoriert.
     */
    private boolean acceptEvents = false;

    private List<WollMuxEvent> eventQueue = new LinkedList<WollMuxEvent>();

    private static EventProcessor singletonInstance;

    private static Thread eventProcessorThread;

    private static EventProcessor getInstance()
    {
      if (singletonInstance == null)
      {
        singletonInstance = new EventProcessor();
        singletonInstance.start();
      }
      return singletonInstance;
    }

    /**
     * Mit dieser Methode ist es möglich die Entgegennahme von Events zu blockieren.
     * Alle eingehenden Events werden ignoriert, wenn accept auf false gesetzt ist
     * und entgegengenommen, wenn accept auf true gesetzt ist.
     * 
     * @param accept
     */
    private void setAcceptEvents(boolean accept)
    {
      acceptEvents = accept;
      if (accept)
        Logger.debug(L.m("EventProcessor: akzeptiere neue Events."));
      else
        Logger.debug(L.m("EventProcessor: blockiere Entgegennahme von Events!"));
    }

    private EventProcessor()
    {
      // starte den eventProcessorThread
      eventProcessorThread = new Thread(new Runnable()
      {
        @Override
        public void run()
        {
          Logger.debug(L.m("Starte EventProcessor-Thread"));
          try
          {
            while (true)
            {
              WollMuxEvent event;
              synchronized (eventQueue)
              {
                while (eventQueue.isEmpty())
                  eventQueue.wait();
                event = eventQueue.remove(0);
              }

              event.process();
            }
          }
          catch (InterruptedException e)
          {
            Logger.error(L.m("EventProcessor-Thread wurde unterbrochen:"));
            Logger.error(e);
          }
          Logger.debug(L.m("Beende EventProcessor-Thread"));
        }
      });
    }

    /**
     * Startet den {@link #eventProcessorThread}.
     * 
     * @author Matthias Benkmann (D-III-ITD-D101)
     */
    private void start()
    {
      eventProcessorThread.start();
    }

    /**
     * Diese Methode fügt ein Event an die eventQueue an wenn der WollMux erfolgreich
     * initialisiert wurde und damit events akzeptieren darf. Anschliessend weckt sie
     * den EventProcessor-Thread.
     * 
     * @param event
     */
    private void addEvent(WollMuxEventHandler.WollMuxEvent event)
    {
      if (acceptEvents) synchronized (eventQueue)
      {
        eventQueue.add(event);
        eventQueue.notifyAll();
      }
    }
  }

  /**
   * Interface für die Events, die dieser EventHandler abarbeitet.
   */
  public interface WollMuxEvent
  {
    /**
     * Startet die Ausführung des Events und darf nur aus dem EventProcessor
     * aufgerufen werden.
     */
    public void process();
  }

  public static class CantStartDialogException extends WollMuxFehlerException
  {
    private static final long serialVersionUID = -1130975078605219254L;

    public CantStartDialogException(java.lang.Exception e)
    {
      super(
        L.m("Der Dialog konnte nicht gestartet werden!\n\nBitte kontaktieren Sie Ihre Systemadministration."),
        e);
    }
  }

  /**
   * Dient als Basisklasse für konkrete Event-Implementierungen.
   */
  private static class BasicEvent implements WollMuxEvent
  {

    private boolean[] lock = new boolean[] { true };

    /**
     * Diese Method ist für die Ausführung des Events zuständig. Nach der Bearbeitung
     * entscheidet der Rückgabewert ob unmittelbar die Bearbeitung des nächsten
     * Events gestartet werden soll oder ob das GUI blockiert werden soll bis das
     * nächste actionPerformed-Event beim EventProcessor eintrifft.
     */
    @Override
    public void process()
    {
      Logger.debug("Process WollMuxEvent " + this.toString());
      try
      {
        doit();
      }
      catch (WollMuxFehlerException e)
      {
        // hier wäre ein showNoConfigInfo möglich - ist aber nicht eindeutig auf no config zurückzuführen
        errorMessage(e);
      }
      // Notnagel für alle Runtime-Exceptions.
      catch (Throwable t)
      {
        Logger.error(t);
      }
    }

    /**
     * Logged die übergebene Fehlermeldung nach Logger.error() und erzeugt ein
     * Dialogfenster mit der Fehlernachricht.
     */
    protected void errorMessage(Throwable t)
    {
      Logger.error(t);
      String msg = "";
      if (t.getMessage() != null) msg += t.getMessage();
      Throwable c = t.getCause();
      /*
       * Bei RuntimeExceptions keine Benutzersichtbare Meldung, weil
       * 
       * 1. der Benutzer damit eh nix anfangen kann
       * 
       * 2. dies typischerweise passiert, wenn der Benutzer das Dokument geschlossen
       * hat, bevor der WollMux fertig war. In diesem Fall will er nicht mit einer
       * Meldung belästigt werden.
       */
      if (c instanceof RuntimeException) return;

      if (c != null)
      {
        msg += "\n\n" + c;
      }
      ModalDialogs.showInfoModal(L.m("WollMux-Fehler"), msg);
    }

    /**
     * Jede abgeleitete Event-Klasse sollte die Methode doit redefinieren, in der die
     * eigentlich event-Bearbeitung erfolgt. Die Methode doit muss alle auftretenden
     * Exceptions selbst behandeln, Fehler die jedoch benutzersichtbar in einem
     * Dialog angezeigt werden sollen, können über eine WollMuxFehlerException nach
     * oben weitergereicht werden.
     */
    protected void doit() throws WollMuxFehlerException
    {};

    /**
     * Diese Methode kann am Ende einer doit()-Methode aufgerufen werden und versucht
     * die Absturzwahrscheinlichkeit von OOo/WollMux zu senken in dem es den
     * GarbageCollector der JavaVM triggert freien Speicher freizugeben. Durch
     * derartige Aufräumaktionen insbesondere nach der Bearbeitung von Events, die
     * viel mit Dokumenten/Cursorn/Uno-Objekten interagieren, wird die Stabilität des
     * WollMux spürbar gesteigert.
     * 
     * In der Vergangenheit gab es z.B. sporadische, nicht immer reproduzierbare
     * Abstürze von OOo, die vermutlich in einem fehlerhaften Speichermanagement in
     * der schwer zu durchschauenden Kette JVM->UNO-Proxies->OOo begründet waren.
     */
    protected void stabilize()
    {
      System.gc();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName();
    }

    /**
     * Setzt den Enable-Status aller OOo-Fenster, die der desktop aktuell liefert auf
     * enabled. Über den Status kann gesteuert werden, ob das Fenster
     * Benutzerinteraktionen wie z.B. Mausklicks auf Menüpunkte oder Tastendrücke
     * verarbeitet. Die Verarbeitung findet nicht statt, wenn enabled==false gesetzt
     * ist, ansonsten schon.
     * 
     * @param enabled
     */
    static void enableAllOOoWindows(boolean enabled)
    {
      try
      {
        XFrames frames = UNO.XFramesSupplier(UNO.desktop).getFrames();
        for (int i = 0; i < frames.getCount(); i++)
        {
          try
          {
            XFrame frame = UNO.XFrame(frames.getByIndex(i));
            XWindow contWin = frame.getContainerWindow();
            if (contWin != null) contWin.setEnable(enabled);
          }
          catch (java.lang.Exception e)
          {
            Logger.error(e);
          }
        }
      }
      catch (java.lang.Exception e)
      {
        Logger.error(e);
      }
    }

    /**
     * Setzt einen lock, der in Verbindung mit setUnlock und der
     * waitForUnlock-Methode verwendet werden kann, um quasi Modalität für nicht
     * modale Dialoge zu realisieren und setzt alle OOo-Fenster auf enabled==false.
     * setLock() sollte stets vor dem Aufruf des nicht modalen Dialogs erfolgen, nach
     * dem Aufruf des nicht modalen Dialogs folgt der Aufruf der
     * waitForUnlock()-Methode. Der nicht modale Dialog erzeugt bei der Beendigung
     * ein ActionEvent, das dafür sorgt, dass setUnlock aufgerufen wird.
     */
    protected void setLock()
    {
      enableAllOOoWindows(false);
      synchronized (lock)
      {
        lock[0] = true;
      }
    }

    /**
     * Macht einen mit setLock() gesetzten Lock rückgängig, und setzt alle
     * OOo-Fenster auf enabled==true und bricht damit eine evtl. wartende
     * waitForUnlock()-Methode ab.
     */
    protected void setUnlock()
    {
      synchronized (lock)
      {
        lock[0] = false;
        lock.notifyAll();
      }
      enableAllOOoWindows(true);
    }

    /**
     * Wartet so lange, bis der vorher mit setLock() gesetzt lock mit der Methode
     * setUnlock() aufgehoben wird. So kann die quasi Modalität nicht modale Dialoge
     * zu realisiert werden. setLock() sollte stets vor dem Aufruf des nicht modalen
     * Dialogs erfolgen, nach dem Aufruf des nicht modalen Dialogs folgt der Aufruf
     * der waitForUnlock()-Methode. Der nicht modale Dialog erzeugt bei der
     * Beendigung ein ActionEvent, das dafür sorgt, dass setUnlock aufgerufen wird.
     */
    protected void waitForUnlock()
    {
      try
      {
        synchronized (lock)
        {
          while (lock[0] == true)
            lock.wait();
        }
      }
      catch (InterruptedException e)
      {}
    }

    /**
     * Dieser ActionListener kann nicht modalen Dialogen übergeben werden und sorgt
     * in Verbindung mit den Methoden setLock() und waitForUnlock() dafür, dass quasi
     * modale Dialoge realisiert werden können.
     */
    protected UnlockActionListener unlockActionListener = new UnlockActionListener();

    protected class UnlockActionListener implements ActionListener
    {
      public ActionEvent actionEvent = null;

      @Override
      public void actionPerformed(ActionEvent arg0)
      {
        actionEvent = arg0;
        setUnlock();
      }
    }
  }

  /**
   * Stellt das WollMuxEvent event in die EventQueue des EventProcessors.
   * 
   * @param event
   */
  private static void handle(WollMuxEvent event)
  {
    EventProcessor.getInstance().addEvent(event);
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das den Dialog AbsenderAuswaehlen startet.
   */
  public static void handleShowDialogAbsenderAuswaehlen()
  {
    handle(new OnShowDialogAbsenderAuswaehlen());
  }

  /**
   * Dieses Event wird vom WollMux-Service (...comp.WollMux) und aus dem
   * WollMuxEventHandler ausgelöst und sorgt dafür, dass der Dialog AbsenderAuswählen
   * gestartet wird.
   * 
   * @author christoph.lutz
   */
  private static class OnShowDialogAbsenderAuswaehlen extends BasicEvent
  {
    @Override
    protected void doit() throws WollMuxFehlerException
    {
      ConfigThingy conf = WollMuxFiles.getWollmuxConf();

      try
      {
        // Konfiguration auslesen:
        ConfigThingy whoAmIconf = requireLastSection(conf, "AbsenderAuswaehlen");
        ConfigThingy PALconf = requireLastSection(conf, "PersoenlicheAbsenderliste");
        ConfigThingy ADBconf = requireLastSection(conf, "AbsenderdatenBearbeiten");

        // Dialog modal starten:
        setLock();
        new AbsenderAuswaehlen(whoAmIconf, PALconf, ADBconf,
          WollMuxFiles.getDatasourceJoiner(), unlockActionListener);
        waitForUnlock();
      }
      catch (Exception e)
      {
        throw new CantStartDialogException(e);
      }

      WollMuxEventHandler.handlePALChangedNotify();
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das den Dialog
   * PersoenlichtAbsenderListe-Verwalten startet.
   */
  public static void handleShowDialogPersoenlicheAbsenderliste()
  {
    handle(new OnShowDialogPersoenlicheAbsenderlisteVerwalten());
  }

  /**
   * Dieses Event wird vom WollMux-Service (...comp.WollMux) und aus dem
   * WollMuxEventHandler ausgelöst und sorgt dafür, dass der Dialog
   * PersönlicheAbsendeliste-Verwalten gestartet wird.
   * 
   * @author christoph.lutz
   */
  private static class OnShowDialogPersoenlicheAbsenderlisteVerwalten extends
      BasicEvent
  {
    @Override
    protected void doit() throws WollMuxFehlerException
    {
      ConfigThingy conf = WollMuxFiles.getWollmuxConf();

      try
      {
        // Konfiguration auslesen:
        ConfigThingy PALconf = requireLastSection(conf, "PersoenlicheAbsenderliste");
        ConfigThingy ADBconf = requireLastSection(conf, "AbsenderdatenBearbeiten");

        // Dialog modal starten:
        setLock();
        new PersoenlicheAbsenderlisteVerwalten(PALconf, ADBconf,
          WollMuxFiles.getDatasourceJoiner(), unlockActionListener);
        waitForUnlock();
}
      catch (Exception e)
      {
        throw new CantStartDialogException(e);
      }

      handlePALChangedNotify();
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das den Funktionsdialog dialogName aufruft und
   * die zurückgelieferten Werte in die entsprechenden FormField-Objekte des
   * Dokuments doc einträgt.
   * 
   * Dieses Event wird vom WollMux-Service (...comp.WollMux) und aus dem
   * WollMuxEventHandler ausgelöst.
   */
  public static void handleFunctionDialog(TextDocumentModel model, String dialogName)
  {
    handle(new OnFunctionDialog(model, dialogName));
  }

  private static class OnFunctionDialog extends BasicEvent
  {
    private TextDocumentModel model;

    private String dialogName;

    private OnFunctionDialog(TextDocumentModel model, String dialogName)
    {
      this.model = model;
      this.dialogName = dialogName;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      // Dialog aus Funktionsdialog-Bibliothek holen:
      Dialog dialog = GlobalFunctions.getInstance().getFunctionDialogs().get(dialogName);
      if (dialog == null)
        throw new WollMuxFehlerException(L.m(
          "Funktionsdialog '%1' ist nicht definiert.", dialogName));

      // Dialoginstanz erzeugen und modal anzeigen:
      Dialog dialogInst = null;
      try
      {
        dialogInst = dialog.instanceFor(new HashMap<Object, Object>());

        setLock();
        dialogInst.show(unlockActionListener, model.getFunctionLibrary(),
          model.getDialogLibrary());
        waitForUnlock();
      }
      catch (ConfigurationErrorException e)
      {
        throw new CantStartDialogException(e);
      }

      // Abbruch, wenn der Dialog nicht mit OK beendet wurde.
      String cmd = unlockActionListener.actionEvent.getActionCommand();
      if (!cmd.equalsIgnoreCase("select")) return;

      // Dem Dokument den Fokus geben, damit die Änderungen des Benutzers
      // transparent mit verfolgt werden können.
      try
      {
        model.getFrame().getContainerWindow().setFocus();
      }
      catch (java.lang.Exception e)
      {
        // keine Gefährdung des Ablaufs falls das nicht klappt.
      }

      // Alle Werte die der Funktionsdialog sicher zurück liefert werden in
      // das Dokument übernommen.
      Collection<String> schema = dialogInst.getSchema();
      Iterator<String> iter = schema.iterator();
      while (iter.hasNext())
      {
        String id = iter.next();
        String value = dialogInst.getData(id).toString();

        // Formularwerte ins Model uns ins Dokument übernehmen.
        model.setFormFieldValue(id, value);
        model.updateFormFields(id);
      }

      stabilize();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ", '" + dialogName
        + "')";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das einen modalen Dialog anzeigt, der wichtige
   * Versionsinformationen über den WollMux, die Konfiguration und die WollMuxBar
   * (nur falls wollmuxBarVersion nicht der Leersting ist) enthält. Anmerkung: das
   * WollMux-Modul hat keine Ahnung, welche WollMuxBar verwendet wird. Daher ist es
   * möglich, über den Parameter wollMuxBarVersion eine Versionsnummer der WollMuxBar
   * zu übergeben, die im Dialog angezeigt wird, falls wollMuxBarVersion nicht der
   * Leerstring ist.
   * 
   * Dieses Event wird vom WollMux-Service (...comp.WollMux) ausgelöst, wenn die
   * WollMux-url "wollmux:about" aufgerufen wurde.
   */
  public static void handleAbout(String wollMuxBarVersion)
  {
    handle(new OnAbout(wollMuxBarVersion));
  }

  private static class OnAbout extends BasicEvent
  {
    private String wollMuxBarVersion;

    private final URL WM_URL = this.getClass().getClassLoader().getResource(
      "data/wollmux_klein.jpg");

    private OnAbout(String wollMuxBarVersion)
    {
      this.wollMuxBarVersion = wollMuxBarVersion;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      Common.setLookAndFeelOnce();

      // non-modal dialog. Set 3rd param to true to make modal
      final JDialog dialog =
        new JDialog((Frame) null, L.m("Info über Vorlagen und Formulare (WollMux)"),
          false);
      JPanel myPanel = new JPanel();
      myPanel.setLayout(new BoxLayout(myPanel, BoxLayout.Y_AXIS));
      myPanel.setBackground(Color.WHITE);
      dialog.setBackground(Color.WHITE);
      dialog.setContentPane(myPanel);
      JPanel imagePanel = new JPanel(new BorderLayout());
      imagePanel.add(new JLabel(new ImageIcon(WM_URL)), BorderLayout.CENTER);
      imagePanel.setOpaque(false);
      Box copyrightPanel = Box.createVerticalBox();
      copyrightPanel.setOpaque(false);
      Box hbox = Box.createHorizontalBox();
      hbox.add(imagePanel);
      hbox.add(copyrightPanel);
      myPanel.add(hbox);

      copyrightPanel.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));
      JLabel label = new JLabel(L.m("WollMux") + " " + WollMuxSingleton.getVersion());
      Font largeFont = label.getFont().deriveFont(15.0f);
      label.setFont(largeFont);
      copyrightPanel.add(label);
      label = new JLabel(L.m("Copyright (c) 2005-2016 Landeshauptstadt München"));
      label.setFont(largeFont);
      copyrightPanel.add(label);
      label = new JLabel(L.m("Lizenz: %1", "European Union Public License"));
      label.setFont(largeFont);
      copyrightPanel.add(label);
      label = new JLabel(L.m("Homepage: %1", "www.wollmux.org"));
      label.setFont(largeFont);
      copyrightPanel.add(label);

      Box authorOuterPanel = Box.createHorizontalBox();
      authorOuterPanel.add(Box.createHorizontalStrut(8));

      Box authorPanelHBox = Box.createHorizontalBox();
      authorPanelHBox.setBorder(BorderFactory.createTitledBorder(L.m("Autoren (sortiert nach LOC)")));
      authorOuterPanel.add(authorPanelHBox);
      authorOuterPanel.add(Box.createHorizontalStrut(8));

      Box authorPanel = Box.createVerticalBox();
      authorPanelHBox.add(authorPanel);
      authorPanel.setOpaque(false);
      myPanel.add(authorOuterPanel);

      int authorSpacing = 1;

      hbox = Box.createHorizontalBox();
      // hbox.add(new JLabel(new ImageIcon(MB_URL)));
      // hbox.add(Box.createHorizontalStrut(8));
      hbox.add(new JLabel("Matthias S. Benkmann"));
      hbox.add(Box.createHorizontalGlue());
      authorPanel.add(hbox);

      hbox = Box.createHorizontalBox();
      // hbox.add(new JLabel(new ImageIcon(CL_URL)));
      // hbox.add(Box.createHorizontalStrut(8));
      hbox.add(new JLabel("Christoph Lutz"));
      hbox.add(Box.createHorizontalGlue());
      authorPanel.add(Box.createVerticalStrut(authorSpacing));
      authorPanel.add(hbox);

      authorPanel.add(Box.createVerticalGlue());

      // Box authorPanel2 = Box.createVerticalBox();
      // authorPanelHBox.add(authorPanel2);
      Box authorPanel2 = authorPanel;

      hbox = Box.createHorizontalBox();
      hbox.add(new JLabel("Daniel Benkmann"));
      hbox.add(Box.createHorizontalGlue());
      authorPanel2.add(Box.createVerticalStrut(authorSpacing));
      authorPanel2.add(hbox);
      hbox = Box.createHorizontalBox();
      hbox.add(new JLabel("Bettina Bauer"));
      hbox.add(Box.createHorizontalGlue());
      authorPanel2.add(Box.createVerticalStrut(authorSpacing));
      authorPanel2.add(hbox);
      hbox = Box.createHorizontalBox();
      hbox.add(new JLabel("Andor Ertsey"));
      hbox.add(Box.createHorizontalGlue());
      authorPanel2.add(Box.createVerticalStrut(authorSpacing));
      authorPanel2.add(hbox);
      hbox = Box.createHorizontalBox();
      hbox.add(new JLabel("Max Meier"));
      hbox.add(Box.createHorizontalGlue());
      authorPanel2.add(Box.createVerticalStrut(authorSpacing));
      authorPanel2.add(hbox);

      authorPanel2.add(Box.createVerticalGlue());

      myPanel.add(Box.createVerticalStrut(8));
      Box infoOuterPanel = Box.createHorizontalBox();
      infoOuterPanel.add(Box.createHorizontalStrut(8));
      JPanel infoPanel = new JPanel(new GridLayout(4, 1));
      infoOuterPanel.add(infoPanel);
      myPanel.add(infoOuterPanel);
      infoOuterPanel.add(Box.createHorizontalStrut(8));
      infoPanel.setOpaque(false);
      infoPanel.setBorder(BorderFactory.createTitledBorder(L.m("Info")));

      infoPanel.add(new JLabel(L.m("WollMux") + " " + WollMuxSingleton.getBuildInfo()));

      if (wollMuxBarVersion != null && !wollMuxBarVersion.equals(""))
        infoPanel.add(new JLabel(L.m("WollMux-Leiste") + " " + wollMuxBarVersion));

      infoPanel.add(new JLabel(L.m("WollMux-Konfiguration:") + " "
        + WollMuxSingleton.getInstance().getConfVersionInfo()));

      infoPanel.add(new JLabel("DEFAULT_CONTEXT: "
        + WollMuxFiles.getDEFAULT_CONTEXT().toExternalForm()));

      myPanel.add(Box.createVerticalStrut(4));

      hbox = Box.createHorizontalBox();
      hbox.add(Box.createHorizontalGlue());
      hbox.add(new JButton(new AbstractAction(L.m("OK"))
      {
        private static final long serialVersionUID = 4527702807001201116L;

        @Override
        public void actionPerformed(ActionEvent e)
        {
          dialog.dispose();
        }
      }));
      hbox.add(Box.createHorizontalGlue());
      myPanel.add(hbox);

      myPanel.add(Box.createVerticalStrut(4));

      dialog.setBackground(Color.WHITE);
      myPanel.setBackground(Color.WHITE);
      dialog.pack();
      int frameWidth = dialog.getWidth();
      int frameHeight = dialog.getHeight();
      Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
      int x = screenSize.width / 2 - frameWidth / 2;
      int y = screenSize.height / 2 - frameHeight / 2;
      dialog.setLocation(x, y);
      dialog.setAlwaysOnTop(true);
      dialog.setVisible(true);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "()";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das den FormularMax4000 aufruft für das Dokument
   * doc.
   * 
   * Dieses Event wird vom WollMux-Service (...comp.WollMux) und aus dem
   * WollMuxEventHandler ausgelöst.
   */
  public static void handleFormularMax4000Show(TextDocumentModel model)
  {
    handle(new OnFormularMax4000Show(model));
  }

  private static class OnFormularMax4000Show extends BasicEvent
  {
    private final TextDocumentModel model;

    private OnFormularMax4000Show(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      // Bestehenden Max in den Vordergrund holen oder neuen Max erzeugen.
      FormularMax4kController max = model.getCurrentFormularMax4000();
      if (max != null)
      {
        max.toFront();
      }
      else
      {
        ActionListener l = new ActionListener()
        {
          @Override
          public void actionPerformed(ActionEvent actionEvent)
          {
            if (actionEvent.getSource() instanceof FormularMax4kController)
              WollMuxEventHandler.handleFormularMax4000Returned(model);
          }
        };

        // Der Konstruktor von FormularMax erwartet hier nur die globalen
        // Funktionsbibliotheken, nicht jedoch die neuen dokumentlokalen
        // Bibliotheken, die das model bereitstellt. Die dokumentlokalen
        // Bibliotheken kann der FM4000 selbst auflösen.
        max =
          new FormularMax4kController(model, l,
            GlobalFunctions.getInstance().getGlobalFunctions(),
            GlobalFunctions.getInstance().getGlobalPrintFunctions());
        model.setCurrentFormularMax4000(max);
        max.run();
      }
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das aufgerufen wird, wenn ein FormularMax4000
   * beendet wird und die entsprechenden internen Referenzen gelöscht werden können.
   * 
   * Dieses Event wird vom EventProcessor geworfen, wenn der FormularMax zurückkehrt.
   */
  public static void handleFormularMax4000Returned(TextDocumentModel model)
  {
    handle(new OnFormularMax4000Returned(model));
  }

  private static class OnFormularMax4000Returned extends BasicEvent
  {
    private TextDocumentModel model;

    private OnFormularMax4000Returned(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      model.setCurrentFormularMax4000(null);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + model.hashCode() + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das aufgerufen wird, wenn ein FormularMax4000
   * beendet wird und die entsprechenden internen Referenzen gelöscht werden können.
   * 
   * Dieses Event wird vom EventProcessor geworfen, wenn der FormularMax zurückkehrt.
   */
  public static void handleMailMergeNewReturned(TextDocumentModel model)
  {
    handle(new OnHandleMailMergeNewReturned(model));
  }

  private static class OnHandleMailMergeNewReturned extends BasicEvent
  {
    private TextDocumentModel model;

    private OnHandleMailMergeNewReturned(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      model.setCurrentMailMergeNew(null);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + model.hashCode() + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das Auskunft darüber gibt, dass ein TextDokument
   * geschlossen wurde und damit auch das TextDocumentModel disposed werden soll.
   * 
   * Dieses Event wird ausgelöst, wenn ein TextDokument geschlossen wird.
   * 
   * @param docInfo
   *          ein {@link DocumentManager.Info} Objekt, an dem das TextDocumentModel
   *          dranhängt des Dokuments, das geschlossen wurde. ACHTUNG! docInfo hat
   *          nicht zwingend ein TextDocumentModel. Es muss
   *          {@link DocumentManager.Info#hasTextDocumentModel()} verwendet werden.
   * 
   * 
   *          ACHTUNG! ACHTUNG! Die Implementierung wurde extra so gewählt, dass hier
   *          ein DocumentManager.Info anstatt direkt eines TextDocumentModel
   *          übergeben wird. Es kam nämlich bei einem Dokument, das schnell geöffnet
   *          und gleich wieder geschlossen wurde zu folgendem Deadlock:
   * 
   *          {@link OnProcessTextDocument} =>
   *          {@link de.muenchen.allg.itd51.wollmux.DocumentManager.TextDocumentInfo#getTextDocumentModel()}
   *          => {@link TextDocumentModel#TextDocumentModel(XTextDocument)} =>
   *          {@link DispatchProviderAndInterceptor#registerDocumentDispatchInterceptor(XFrame)}
   *          => OOo Proxy =>
   *          {@link GlobalEventListener#notifyEvent(com.sun.star.document.EventObject)}
   *          ("OnUnload") =>
   *          {@link de.muenchen.allg.itd51.wollmux.DocumentManager.TextDocumentInfo#hasTextDocumentModel()}
   * 
   *          Da {@link TextDocumentInfo} synchronized ist kam es zum Deadlock.
   * 
   */
  public static void handleTextDocumentClosed(DocumentManager.Info docInfo)
  {
    handle(new OnTextDocumentClosed(docInfo));
  }

  private static class OnTextDocumentClosed extends BasicEvent
  {
    private DocumentManager.Info docInfo;

    private OnTextDocumentClosed(DocumentManager.Info doc)
    {
      this.docInfo = doc;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      if (docInfo.hasTextDocumentModel()) docInfo.getTextDocumentModel().dispose();
    }

    @Override
    public String toString()
    {
      String code = "unknown";
      if (docInfo.hasTextDocumentModel())
        code = "" + docInfo.getTextDocumentModel().hashCode();
      return this.getClass().getSimpleName() + "(#" + code + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das die eigentliche Dokumentbearbeitung eines
   * TextDokuments startet.
   * 
   * @param xTextDoc
   *          Das XTextDocument, das durch den WollMux verarbeitet werden soll.
   * @param visible
   *          false zeigt an, dass das Dokument (bzw. das zugehörige Fenster)
   *          unsichtbar ist.
   */
  public static void handleProcessTextDocument(XTextDocument xTextDoc,
      boolean visible)
  {
    handle(new OnProcessTextDocument(xTextDoc, visible));
  }

  /**
   * Dieses Event wird immer dann ausgelöst, wenn der GlobalEventBroadcaster von OOo
   * ein ON_NEW oder ein ON_LOAD-Event wirft. Das Event sorgt dafür, dass die
   * eigentliche Dokumentbearbeitung durch den WollMux angestossen wird.
   * 
   * @author christoph.lutz
   */
  private static class OnProcessTextDocument extends BasicEvent
  {
    XTextDocument xTextDoc;

    boolean visible;

    public OnProcessTextDocument(XTextDocument xTextDoc, boolean visible)
    {
      this.xTextDoc = xTextDoc;
      this.visible = visible;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      if (xTextDoc == null) return;

      TextDocumentModel model = DocumentManager.getTextDocumentModel(xTextDoc);

      // Konfigurationsabschnitt Textdokument verarbeiten falls Dok sichtbar:
      if (visible)
        try
        {
          ConfigThingy tds =
            WollMuxFiles.getWollmuxConf().query("Fenster").query("Textdokument").getLastChild();
          model.setWindowViewSettings(tds);
        }
        catch (NodeNotFoundException e)
        {}

      // Workaround für OOo-Issue 103137 ggf. anwenden:
      if (Workarounds.applyWorkaroundForOOoIssue103137())
      {
        // Alle mit OOo 2 erstellen Dokumente, die Textstellen enthalten, die von
        // älteren WollMux-Versionen ein- oder ausgeblendet wurden sind potentiell
        // betroffen. Ein- und Ausblendungen in WollMux-Formularen werden
        // glücklicherweise auch ohne Workaround vom WollMux korrigiert, da der
        // WollMux beim Erzeugen der FormularGUI alle Formularwerte und damit auch
        // die Sichtbarkeiten explizit setzt. Damit bleiben nur die Sachleitenden
        // Verfügungen übrig, die vom WollMux bisher noch nicht automatisch
        // korrigiert wurden. Dieser Workaround macht das nun:
        if (model.getPrintFunctions().contains(
          SachleitendeVerfuegung.PRINT_FUNCTION_NAME))
          SachleitendeVerfuegung.workaround103137(model);
      }

      // Mögliche Aktionen für das neu geöffnete Dokument:
      DocumentCommandInterpreter dci = new DocumentCommandInterpreter(model, WollMuxFiles.isDebugMode());

      try
      {
        // Globale Dokumentkommandos wie z.B. setType, setPrintFunction, ...
        // auswerten.
        dci.scanGlobalDocumentCommands();

        int actions =
          model.evaluateDocumentActions(GlobalFunctions.getInstance().getDocumentActionFunctions().iterator());

        // Bei Vorlagen: Ausführung der Dokumentkommandos
        if ((actions < 0 && model.isTemplate()) || (actions == Integer.MAX_VALUE))
        {
          dci.executeTemplateCommands();

          // manche Kommandos sind erst nach der Expansion verfügbar
          dci.scanGlobalDocumentCommands();
        }
        // insertFormValue-Kommandos auswerten
        dci.scanInsertFormValueCommands();

        // Bei Formularen:
        // Anmerkung: actions == allactions wird NICHT so interpretiert, dass auch
        // bei Dokumenten ohne Formularfunktionen der folgende Abschnitt ausgeführt
        // wird. Dies würde unsinnige Ergebnisse verursachen.
        if (actions != 0 && model.isFormDocument())
        {
          // Konfigurationsabschnitt Fenster/Formular verarbeiten falls Dok sichtbar
          if (visible)
            try
            {
              model.setDocumentZoom(WollMuxFiles.getWollmuxConf().query("Fenster").query(
                "Formular").getLastChild().query("ZOOM"));
            }
            catch (java.lang.Exception e)
            {}

          // FormGUI starten, falls es kein Teil eines Multiform-Dokuments ist.
          if (!model.isPartOfMultiformDocument())
          {
            FormModel fm;
            try
            {
              fm = SingleDocumentFormModel.createSingleDocumentFormModel(model, visible);
            }
            catch (InvalidFormDescriptorException e)
            {
              throw new WMCommandsFailedException(
                L.m("Die Vorlage bzw. das Formular enthält keine gültige Formularbeschreibung\n\nBitte kontaktieren Sie Ihre Systemadministration."));
            }

            model.setFormModel(fm);
            fm.startFormGUI();
          }
        }
      }
      catch (java.lang.Exception e)
      {
        throw new WollMuxFehlerException(L.m("Fehler bei der Dokumentbearbeitung."),
          e);
      }

      // Registrierte XEventListener (etwas später) informieren, dass die
      // Dokumentbearbeitung fertig ist.
      handleNotifyDocumentEventListener(null, ON_WOLLMUX_PROCESSING_FINISHED,
        model.doc);

      // ContextChanged auslösen, damit die Dispatches aktualisiert werden.
      try
      {
        model.getFrame().contextChanged();
      }
      catch (java.lang.Exception e)
      {}

      stabilize();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + xTextDoc.hashCode() + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Wird wollmux:Open mit der Option FORMGUIS "merged" gestartet, so werden zuerst
   * die Einzeldokumente geöffnet und dann dieses Event aufgerufen, das dafür
   * zuständig ist, die eine FormGUI für den MultiDocument-Modus zu erzeugen und zu
   * starten.
   * 
   * @param docs
   *          Ein Vector of TextDocumentModels, die in einem Multiformular
   *          zusammengefasst werden sollen.
   */
  public static void handleProcessMultiform(
      Vector /* of TextDocumentModel */<TextDocumentModel> docs,
      ConfigThingy buttonAnpassung)
  {
    handle(new OnProcessMultiform(docs, buttonAnpassung));
  }

  private static class OnProcessMultiform extends BasicEvent
  {
    Vector<TextDocumentModel> docs;

    ConfigThingy buttonAnpassung;

    public OnProcessMultiform(
        Vector /* of TextDocumentModel */<TextDocumentModel> docs,
        ConfigThingy buttonAnpassung)
    {
      this.docs = docs;
      this.buttonAnpassung = buttonAnpassung;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      FormModel fm;
      try
      {
        fm = MultiDocumentFormModel.createMultiDocumentFormModel(docs, buttonAnpassung);
      }
      catch (InvalidFormDescriptorException e)
      {
        throw new WollMuxFehlerException(
          L.m("Fehler bei der Dokumentbearbeitung."),
          new WMCommandsFailedException(
            L.m("Die Vorlage bzw. das Formular enthält keine gültige Formularbeschreibung\n\nBitte kontaktieren Sie Ihre Systemadministration.")));
      }

      // FormModel in allen Dokumenten registrieren:
      for (Iterator<TextDocumentModel> iter = docs.iterator(); iter.hasNext();)
      {
        TextDocumentModel doc = iter.next();
        doc.setFormModel(fm);
      }

      // FormGUI starten:
      fm.startFormGUI();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + docs + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das die Bearbeitung aller neu hinzugekommener
   * Dokumentkommandos übernimmt.
   * 
   * Dieses Event wird immer dann ausgelöst, wenn nach dem Öffnen eines Dokuments
   * neue Dokumentkommandos eingefügt wurden, die nun bearbeitet werden sollen - z.B.
   * wenn ein Textbaustein eingefügt wurde.
   * 
   * @param model
   *          Das TextDocumentModel, das durch den WollMux verarbeitet werden soll.
   */
  public static void handleReprocessTextDocument(TextDocumentModel model)
  {
    handle(new OnReprocessTextDocument(model));
  }

  private static class OnReprocessTextDocument extends BasicEvent
  {
    TextDocumentModel model;

    public OnReprocessTextDocument(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      if (model == null) return;

      // Dokument mit neuen Dokumentkommandos über den
      // DocumentCommandInterpreter bearbeiten:
      model.getDocumentCommands().update();
      DocumentCommandInterpreter dci =
        new DocumentCommandInterpreter(model, WollMuxFiles.isDebugMode());
      try
      {
        dci.executeTemplateCommands();

        // manche Kommandos sind erst nach der Expansion verfügbar
        dci.scanGlobalDocumentCommands();
        dci.scanInsertFormValueCommands();
      }
      catch (java.lang.Exception e)
      {
        // Hier wird keine Exception erwartet, da Fehler (z.B. beim manuellen
        // Einfügen von Textbausteinen) bereits dort als Popup angezeigt werden
        // sollen, wo sie auftreten.
      }

      stabilize();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Obsolete, aber aus Kompatibilitätgründen noch vorhanden. Bitte handleOpen()
   * statt dessen verwenden.
   * 
   * Erzeugt ein neues WollMuxEvent, welches dafür sorgt, dass ein Dokument geöffnet
   * wird.
   * 
   * Dieses Event wird gestartet, wenn der WollMux-Service (...comp.WollMux) das
   * Dispatch-Kommando wollmux:openTemplate bzw. wollmux:openDocument empfängt und
   * sort dafür, dass das entsprechende Dokument geöffnet wird.
   * 
   * @param fragIDs
   *          Eine List mit fragIDs, wobei das erste Element die FRAG_ID des zu
   *          öffnenden Dokuments beinhalten muss. Weitere Elemente werden in eine
   *          Liste zusammengefasst und als Parameter für das Dokumentkommando
   *          insertContent verwendet.
   * @param asTemplate
   *          true, wenn das Dokument als "Unbenannt X" geöffnet werden soll (also im
   *          "Template-Modus") und false, wenn das Dokument zum Bearbeiten geöffnet
   *          werden soll.
   */
  public static void handleOpenDocument(List<String> fragIDs, boolean asTemplate)
  {
    handle(new OnOpenDocument(fragIDs, asTemplate));
  }

  private static class OnOpenDocument extends BasicEvent
  {
    private boolean asTemplate;

    private List<String> fragIDs;

    private OnOpenDocument(List<String> fragIDs, boolean asTemplate)
    {
      this.fragIDs = fragIDs;
      this.asTemplate = asTemplate;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      // Baue ein ConfigThingy (als String), das die neue open-Methode versteht
      // und leite es weiter an diese.
      Iterator<String> iter = fragIDs.iterator();
      StringBuffer fragIdStr = new StringBuffer();
      while (iter.hasNext())
      {
        fragIdStr.append("'");
        fragIdStr.append(iter.next());
        fragIdStr.append("' ");
      }
      handleOpen("AS_TEMPLATE '" + asTemplate
        + "' FORMGUIS 'independent' Fragmente( FRAG_ID_LIST ("
        + fragIdStr.toString() + "))");
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "("
        + ((asTemplate) ? "asTemplate" : "asDocument") + ", " + fragIDs + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Öffnet ein oder mehrere Dokumente anhand der Beschreibung openConfStr
   * (ConfigThingy-Syntax) und ist ausführlicher beschrieben unter
   * http://limux.tvc.muenchen
   * .de/wiki/index.php/Schnittstellen_des_WollMux_f%C3%BCr_Experten#wollmux:Open
   * 
   * Dieses Event wird gestartet, wenn der WollMux-Service (...comp.WollMux) das
   * Dispatch-Kommando wollmux:open empfängt.
   * 
   * @param openConfStr
   *          Die Beschreibung der zu öffnenden Fragmente in ConfigThingy-Syntax
   */
  public static void handleOpen(String openConfStr)
  {
    handle(new OnOpen(openConfStr));
  }

  private static class OnOpen extends BasicEvent
  {
    private String openConfStr;

    private OnOpen(String openConfStr)
    {
      this.openConfStr = openConfStr;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      boolean asTemplate = true;
      boolean merged = false;
      ConfigThingy conf;
      ConfigThingy fragConf;
      ConfigThingy buttonAnpassung = null;
      try
      {
        conf = new ConfigThingy("OPEN", null, new StringReader(openConfStr));
        fragConf = conf.get("Fragmente");
      }
      catch (java.lang.Exception e)
      {
        throw new WollMuxFehlerException(
          L.m("Fehlerhaftes Kommando 'wollmux:Open'"), e);
      }

      try
      {
        asTemplate = conf.get("AS_TEMPLATE", 1).toString().equalsIgnoreCase("true");
      }
      catch (java.lang.Exception x)
      {}

      try
      {
        buttonAnpassung = conf.get("Buttonanpassung", 1);
      }
      catch (java.lang.Exception x)
      {}

      try
      {
        merged = conf.get("FORMGUIS", 1).toString().equalsIgnoreCase("merged");
      }
      catch (java.lang.Exception x)
      {}

      Iterator<ConfigThingy> iter = fragConf.iterator();
      Vector<TextDocumentModel> docs = new Vector<TextDocumentModel>();
      while (iter.hasNext())
      {
        ConfigThingy fragListConf = iter.next();
        List<String> fragIds = new Vector<String>();
        Iterator<ConfigThingy> fragIter = fragListConf.iterator();
        while (fragIter.hasNext())
        {
          fragIds.add(fragIter.next().toString());
        }
        if (!fragIds.isEmpty())
        {
          TextDocumentModel model = openTextDocument(fragIds, asTemplate, merged);
          if (model != null) docs.add(model);
        }
      }

      if (merged)
      {
        handleProcessMultiform(docs, buttonAnpassung);
      }
    }

    /**
     * 
     * @param fragIDs
     * @param asTemplate
     * @param asPartOfMultiform
     * @return
     * @throws WollMuxFehlerException
     */
    private TextDocumentModel openTextDocument(List<String> fragIDs,
        boolean asTemplate, boolean asPartOfMultiform) throws WollMuxFehlerException
    {
      // das erste Argument ist das unmittelbar zu landende Textfragment und
      // wird nach urlStr aufgelöst. Alle weiteren Argumente (falls vorhanden)
      // werden nach argsUrlStr aufgelöst.
      String loadUrlStr = "";
      String[] fragUrls = new String[fragIDs.size() - 1];
      String urlStr = "";

      Iterator<String> iter = fragIDs.iterator();
      for (int i = 0; iter.hasNext(); ++i)
      {
        String frag_id = iter.next();

        // Fragment-URL holen und aufbereiten:
        Vector<String> urls = new Vector<String>();

        java.lang.Exception error =
          new ConfigurationErrorException(L.m(
            "Das Textfragment mit der FRAG_ID '%1' ist nicht definiert!", frag_id));
        try
        {
          urls = VisibleTextFragmentList.getURLsByID(frag_id);
        }
        catch (InvalidIdentifierException e)
        {
          error = e;
        }
        if (urls.size() == 0)
        {
          throw new WollMuxFehlerException(
            L.m(
              "Die URL zum Textfragment mit der FRAG_ID '%1' kann nicht bestimmt werden:",
              frag_id), error);
        }

        // Nur die erste funktionierende URL verwenden. Dazu werden alle URL zu
        // dieser FRAG_ID geprüft und in die Variablen loadUrlStr und fragUrls
        // übernommen.
        String errors = "";
        boolean found = false;
        Iterator<String> iterUrls = urls.iterator();
        while (iterUrls.hasNext() && !found)
        {
          urlStr = iterUrls.next();

          // URL erzeugen und prüfen, ob sie aufgelöst werden kann
          URL url;
          try
          {
            url = WollMuxFiles.makeURL(urlStr);
            urlStr = UNO.getParsedUNOUrl(url.toExternalForm()).Complete;
            url = WollMuxFiles.makeURL(urlStr);
            WollMuxSingleton.checkURL(url);
          }            
          catch (MalformedURLException e)
          {
            Logger.log(e);
            errors +=
              L.m("Die URL '%1' ist ungültig:", urlStr) + "\n"
                + e.getLocalizedMessage() + "\n\n";
            continue;
          }          
          catch (IOException e)
          {
            Logger.log(e);
            errors += e.getLocalizedMessage() + "\n\n";
            continue;
          }

          found = true;
        }

        if (!found)
        {
          throw new WollMuxFehlerException(L.m(
            "Das Textfragment mit der FRAG_ID '%1' kann nicht aufgelöst werden:",
            frag_id)
            + "\n\n" + errors);
        }

        // URL in die in loadUrlStr (zum sofort öffnen) und in argsUrlStr (zum
        // später öffnen) aufnehmen
        if (i == 0)
        {
          loadUrlStr = urlStr;
        }
        else
        {
          fragUrls[i - 1] = urlStr;
        }
      }

      // open document as Template (or as document):
      TextDocumentModel model = null;
      try
      {
        XComponent doc = UNO.loadComponentFromURL(loadUrlStr, asTemplate, true);

        if (UNO.XTextDocument(doc) != null)
        {
          model = DocumentManager.getTextDocumentModel(UNO.XTextDocument(doc));
          model.setFragUrls(fragUrls);
          if (asPartOfMultiform)
            model.setPartOfMultiformDocument(asPartOfMultiform);
        }
      }
      catch (java.lang.Exception x)
      {
        // sollte eigentlich nicht auftreten, da bereits oben geprüft.
        throw new WollMuxFehlerException(L.m(
          "Die Vorlage mit der URL '%1' kann nicht geöffnet werden.", loadUrlStr), x);
      }
      return model;
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "('" + openConfStr + "')";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das dafür sorgt, dass alle registrierten
   * XPALChangeEventListener geupdated werden.
   */
  public static void handlePALChangedNotify()
  {
    handle(new OnPALChangedNotify());
  }

  /**
   * Dieses Event wird immer dann erzeugt, wenn ein Dialog zur Bearbeitung der PAL
   * geschlossen wurde und immer dann wenn die PAL z.B. durch einen
   * wollmux:setSender-Befehl geändert hat. Das Event sorgt dafür, dass alle im
   * WollMuxSingleton registrierten XPALChangeListener geupdatet werden.
   * 
   * @author christoph.lutz
   */
  private static class OnPALChangedNotify extends BasicEvent
  {
    @Override
    protected void doit()
    {
      // registrierte PALChangeListener updaten
      Iterator<XPALChangeEventListener> i = PersoenlicheAbsenderliste.getInstance().iterator();
      while (i.hasNext())
      {
        Logger.debug2("OnPALChangedNotify: Update XPALChangeEventListener");
        EventObject eventObject = new EventObject();
        eventObject.Source = PersoenlicheAbsenderliste.getInstance();
        try
        {
          i.next().updateContent(eventObject);
        }
        catch (java.lang.Exception x)
        {
          i.remove();
        }
      }

      // Cache und LOS auf Platte speichern.
      try
      {
        WollMuxFiles.getDatasourceJoiner().saveCacheAndLOS(WollMuxFiles.getLosCacheFile());
      }
      catch (IOException e)
      {
        Logger.error(e);
      }
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent zum setzen des aktuellen Absenders.
   * 
   * @param senderName
   *          Name des Senders in der Form "Nachname, Vorname (Rolle)" wie sie auch
   *          der PALProvider bereithält.
   * @param idx
   *          der zum Sender senderName passende index in der sortierten Senderliste
   *          - dient zur Konsistenz-Prüfung, damit kein Sender gesetzt wird, wenn
   *          die PAL der setzenden Komponente nicht mit der PAL des WollMux
   *          übereinstimmt.
   */
  public static void handleSetSender(String senderName, int idx)
  {
    handle(new OnSetSender(senderName, idx));
  }

  /**
   * Dieses Event wird ausgelöst, wenn im WollMux-Service die methode setSender
   * aufgerufen wird. Es sort dafür, dass ein neuer Absender gesetzt wird.
   * 
   * @author christoph.lutz
   */
  private static class OnSetSender extends BasicEvent
  {
    private String senderName;

    private int idx;

    public OnSetSender(String senderName, int idx)
    {
      this.senderName = senderName;
      this.idx = idx;
    }

    @Override
    protected void doit()
    {
      String[] pal = PersoenlicheAbsenderliste.getInstance().getPALEntries();

      // nur den neuen Absender setzen, wenn index und sender übereinstimmen,
      // d.h.
      // die Absenderliste der entfernten WollMuxBar konsistent war.
      if (idx >= 0 && idx < pal.length && pal[idx].toString().equals(senderName))
      {
        DJDatasetListElement[] palDatasets =
            PersoenlicheAbsenderliste.getInstance().getSortedPALEntries();
        palDatasets[idx].getDataset().select();
      }
      else
      {
        Logger.error(L.m(
          "Setzen des Senders '%1' schlug fehl, da der index '%2' nicht mit der PAL übereinstimmt (Inkonsistenzen?)",
          senderName, idx));
      }

      WollMuxEventHandler.handlePALChangedNotify();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + senderName + ", " + idx + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, welches dafür sorgt, dass alle Formularfelder
   * Dokument auf den neuen Wert gesetzt werden. Bei Formularfeldern mit
   * TRAFO-Funktion wird die Transformation entsprechend durchgeführt.
   * 
   * Dieses Event wird (derzeit) vom FormModelImpl ausgelöst, wenn in der
   * Formular-GUI der Wert des Formularfeldes fieldID geändert wurde und sorgt dafür,
   * dass die Wertänderung auf alle betroffenen Formularfelder im Dokument doc
   * übertragen werden.
   * 
   * @param idToFormValues
   *          Eine HashMap die unter dem Schlüssel fieldID den Vektor aller
   *          FormFields mit der ID fieldID liefert.
   * @param fieldId
   *          Die ID der Formularfelder, deren Werte angepasst werden sollen.
   * @param newValue
   *          Der neue untransformierte Wert des Formularfeldes.
   * @param funcLib
   *          Die Funktionsbibliothek, die zur Gewinnung der Trafo-Funktion verwendet
   *          werden soll.
   */
  public static void handleFormValueChanged(TextDocumentModel model, String fieldId,
      String newValue)
  {
    handle(new OnFormValueChanged(model, fieldId, newValue));
  }

  private static class OnFormValueChanged extends BasicEvent
  {
    private String fieldId;

    private String newValue;

    private TextDocumentModel model;

    public OnFormValueChanged(TextDocumentModel model, String fieldId,
        String newValue)
    {
      this.fieldId = fieldId;
      this.newValue = newValue;
      this.model = model;
    }

    @Override
    protected void doit()
    {
      model.setFormFieldValue(fieldId, newValue);
      model.updateFormFields(fieldId);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + fieldId + "', '" + newValue
        + "')";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, welches dafür sorgt, dass alle
   * Sichtbarkeitselemente (Dokumentkommandos oder Bereiche mit Namensanhang 'GROUPS
   * ...') im übergebenen Dokument, die einer bestimmten Gruppe groupId zugehören
   * ein- oder ausgeblendet werden.
   * 
   * Dieses Event wird (derzeit) vom FormModelImpl ausgelöst, wenn in der
   * Formular-GUI bestimmte Text-Teile des übergebenen Dokuments ein- oder
   * ausgeblendet werden sollen. Auch das PrintModel verwendet dieses Event, wenn
   * XPrintModel.setGroupVisible() aufgerufen wurde.
   * 
   * @param model
   *          Das TextDocumentModel, welches die Sichtbarkeitselemente enthält.
   * @param groupId
   *          Die GROUP (ID) der ein/auszublendenden Gruppe.
   * @param visible
   *          Der neue Sichtbarkeitsstatus (true=sichtbar, false=ausgeblendet)
   * @param listener
   *          Der listener, der nach Durchführung des Events benachrichtigt wird
   *          (kann auch null sein, dann gibt's keine Nachricht).
   * 
   * @author Christoph Lutz (D-III-ITD-5.1)
   */
  public static void handleSetVisibleState(TextDocumentModel model, String groupId,
      boolean visible, ActionListener listener)
  {
    handle(new OnSetVisibleState(model, groupId, visible, listener));
  }

  private static class OnSetVisibleState extends BasicEvent
  {
    private TextDocumentModel model;

    private String groupId;

    private boolean visible;

    private ActionListener listener;

    public OnSetVisibleState(TextDocumentModel model, String groupId,
        boolean visible, ActionListener listener)
    {
      this.model = model;
      this.groupId = groupId;
      this.visible = visible;
      this.listener = listener;
    }

    @Override
    protected void doit()
    {
      try
      {
        // invisibleGroups anpassen:
        HashSet<String> invisibleGroups = model.getInvisibleGroups();
        if (visible)
          invisibleGroups.remove(groupId);
        else
          invisibleGroups.add(groupId);

        VisibilityElement firstChangedElement = null;

        // Sichtbarkeitselemente durchlaufen und alle ggf. updaten:
        Iterator<VisibilityElement> iter = model.visibleElementsIterator();
        while (iter.hasNext())
        {
          VisibilityElement visibleElement = iter.next();
          Set<String> groups = visibleElement.getGroups();
          if (!groups.contains(groupId)) continue;

          // Visibility-Status neu bestimmen:
          boolean setVisible = true;
          Iterator<String> i = groups.iterator();
          while (i.hasNext())
          {
            String groupId = i.next();
            if (invisibleGroups.contains(groupId)) setVisible = false;
          }

          // Element merken, dessen Sichtbarkeitsstatus sich zuerst ändert und
          // den focus (ViewCursor) auf den Start des Bereichs setzen. Da das
          // Setzen eines ViewCursors in einen unsichtbaren Bereich nicht
          // funktioniert, wird die Methode focusRangeStart zwei mal aufgerufen,
          // je nach dem, ob der Bereich vor oder nach dem Setzen des neuen
          // Sichtbarkeitsstatus sichtbar ist.
          if (setVisible != visibleElement.isVisible()
            && firstChangedElement == null)
          {
            firstChangedElement = visibleElement;
            if (firstChangedElement.isVisible()) focusRangeStart(visibleElement);
          }

          // neuen Sichtbarkeitsstatus setzen:
          try
          {
            visibleElement.setVisible(setVisible);
          }
          catch (RuntimeException e)
          {
            // Absicherung gegen das manuelle Löschen von Dokumentinhalten
          }
        }

        // Den Cursor (nochmal) auf den Anfang des Ankers des Elements setzen,
        // dessen Sichtbarkeitsstatus sich zuerst geändert hat (siehe Begründung
        // oben).
        if (firstChangedElement != null && firstChangedElement.isVisible())
          focusRangeStart(firstChangedElement);
      }
      catch (java.lang.Exception e)
      {
        Logger.error(e);
      }

      if (listener != null) listener.actionPerformed(null);
    }

    /**
     * Diese Methode setzt den ViewCursor auf den Anfang des Ankers des
     * Sichtbarkeitselements.
     * 
     * @param visibleElement
     *          Das Sichtbarkeitselement, auf dessen Anfang des Ankers der ViewCursor
     *          gesetzt werden soll.
     */
    private void focusRangeStart(VisibilityElement visibleElement)
    {
      try
      {
        model.getViewCursor().gotoRange(visibleElement.getAnchor().getStart(), false);
      }
      catch (java.lang.Exception e)
      {}
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "('" + groupId + "', " + visible
        + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein Event, das den ViewCursor des Dokuments auf das aktuell in der
   * Formular-GUI bearbeitete Formularfeld setzt.
   * 
   * Dieses Event wird (derzeit) vom FormModelImpl ausgelöst, wenn in der
   * Formular-GUI ein Formularfeld den Fokus bekommen hat und es sorgt dafür, dass
   * der View-Cursor des Dokuments das entsprechende FormField im Dokument anspringt.
   * 
   * @param idToFormValues
   *          Eine HashMap die unter dem Schlüssel fieldID den Vektor aller
   *          FormFields mit der ID fieldID liefert.
   * @param fieldId
   *          die ID des Formularfeldes das den Fokus bekommen soll. Besitzen mehrere
   *          Formularfelder diese ID, so wird bevorzugt das erste Formularfeld aus
   *          dem Vektor genommen, das keine Trafo enthält. Ansonsten wird das erste
   *          Formularfeld im Vektor verwendet.
   */
  public static void handleFocusFormField(TextDocumentModel model, String fieldId)
  {
    handle(new OnFocusFormField(model, fieldId));
  }

  private static class OnFocusFormField extends BasicEvent
  {
    private TextDocumentModel model;

    private String fieldId;

    public OnFocusFormField(TextDocumentModel model, String fieldId)
    {
      this.model = model;
      this.fieldId = fieldId;
    }

    @Override
    protected void doit()
    {
      model.focusFormField(fieldId);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + model.doc + ", '" + fieldId
        + "')";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein Event, das die Position und Größe des übergebenen Dokument-Fensters
   * auf die vorgegebenen Werte setzt. ACHTUNG: Die Maßangaben beziehen sich auf die
   * linke obere Ecke des Fensterinhalts OHNE die Titelzeile und die
   * Fensterdekoration des Rahmens. Um die linke obere Ecke des gesamten Fensters
   * richtig zu setzen, müssen die Größenangaben des Randes der Fensterdekoration und
   * die Höhe der Titelzeile VOR dem Aufruf der Methode entsprechend eingerechnet
   * werden.
   * 
   * @param model
   *          Das XModel-Interface des Dokuments dessen Position/Größe gesetzt werden
   *          soll.
   * @param docX
   *          Die linke obere Ecke des Fensterinhalts X-Koordinate der Position in
   *          Pixel, gezählt von links oben.
   * @param docY
   *          Die Y-Koordinate der Position in Pixel, gezählt von links oben.
   * @param docWidth
   *          Die Größe des Dokuments auf der X-Achse in Pixel
   * @param docHeight
   *          Die Größe des Dokuments auf der Y-Achse in Pixel. Auch hier wird die
   *          Titelzeile des Rahmens nicht beachtet und muss vorher entsprechend
   *          eingerechnet werden.
   */
  public static void handleSetWindowPosSize(TextDocumentModel model, int docX,
      int docY, int docWidth, int docHeight)
  {
    handle(new OnSetWindowPosSize(model, docX, docY, docWidth, docHeight));
  }

  /**
   * Dieses Event wird vom FormModelImpl ausgelöst, wenn die Formular-GUI die
   * Position und die Ausmasse des Dokuments verändert. Ruft direkt setWindowsPosSize
   * der UNO-API auf.
   * 
   * @author christoph.lutz
   */
  private static class OnSetWindowPosSize extends BasicEvent
  {
    private TextDocumentModel model;

    private int docX, docY, docWidth, docHeight;

    public OnSetWindowPosSize(TextDocumentModel model, int docX, int docY,
        int docWidth, int docHeight)
    {
      this.model = model;
      this.docX = docX;
      this.docY = docY;
      this.docWidth = docWidth;
      this.docHeight = docHeight;
    }

    @Override
    protected void doit()
    {
      model.setWindowPosSize(docX, docY, docWidth, docHeight);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + docX + ", " + docY + ", "
        + docWidth + ", " + docHeight + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein Event, das die Anzeige des übergebenen Dokuments auf sichtbar oder
   * unsichtbar schaltet. Dabei wird direkt die entsprechende Funktion der UNO-API
   * verwendet.
   * 
   * @param model
   *          Das XModel interface des dokuments, welches sichtbar oder unsichtbar
   *          geschaltet werden soll.
   * @param visible
   *          true, wenn das Dokument sichtbar geschaltet werden soll und false, wenn
   *          das Dokument unsichtbar geschaltet werden soll.
   */
  public static void handleSetWindowVisible(TextDocumentModel model, boolean visible)
  {
    handle(new OnSetWindowVisible(model, visible));
  }

  /**
   * Dieses Event wird vom FormModelImpl ausgelöst, wenn die Formular-GUI das
   * bearbeitete Dokument sichtbar/unsichtbar schalten möchte. Ruft direkt setVisible
   * der UNO-API auf.
   * 
   * @author christoph.lutz
   */
  private static class OnSetWindowVisible extends BasicEvent
  {
    private TextDocumentModel model;

    boolean visible;

    public OnSetWindowVisible(TextDocumentModel model, boolean visible)
    {
      this.model = model;
      this.visible = visible;
    }

    @Override
    protected void doit()
    {
      model.setWindowVisible(visible);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + visible + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein Event, das das übergebene Dokument schließt.
   * 
   * @param model
   *          Das zu schließende TextDocumentModel.
   */
  public static void handleCloseTextDocument(TextDocumentModel model)
  {
    handle(new OnCloseTextDocument(model));
  }

  /**
   * Dieses Event wird vom FormModelImpl ausgelöst, wenn der Benutzer die
   * Formular-GUI schließt und damit auch das zugehörige TextDokument geschlossen
   * werden soll.
   * 
   * @author christoph.lutz
   */
  private static class OnCloseTextDocument extends BasicEvent
  {
    private TextDocumentModel model;

    public OnCloseTextDocument(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit()
    {
      model.close();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + model.hashCode() + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein Event, das das übergebene Dokument in eine temporäre Datei
   * speichert, eine externe Anwendung mit dieser aufruft und das Dokument dann
   * schließt, wobei der ExterneAnwendungen-Abschnitt zu ext die näheren Details wie
   * den FILTER regelt.
   * 
   * @param model
   *          Das an die externe Anwendung weiterzureichende TextDocumentModel.
   * @param ext
   *          identifiziert den entsprechenden Eintrag im Abschnitt
   *          ExterneAnwendungen.
   */
  public static void handleCloseAndOpenExt(TextDocumentModel model, String ext)
  {
    handle(new OnCloseAndOpenExt(model, ext));
  }

  /**
   * Dieses Event wird vom FormModelImpl ausgelöst, wenn der Benutzer die Aktion
   * "closeAndOpenExt" aktiviert hat.
   * 
   * @author matthias.benkmann
   */
  private static class OnCloseAndOpenExt extends BasicEvent
  {
    private TextDocumentModel model;

    private String ext;

    public OnCloseAndOpenExt(TextDocumentModel model, String ext)
    {
      this.model = model;
      this.ext = ext;
    }

    @Override
    protected void doit()
    {
      try
      {
        OpenExt openExt = new OpenExt(ext, WollMuxFiles.getWollmuxConf());
        openExt.setSource(UNO.XStorable(model.doc));
        openExt.storeIfNecessary();
        openExt.launch(new OpenExt.ExceptionHandler()
        {
          @Override
          public void handle(Exception x)
          {
            Logger.error(x);
          }
        });
      }
      catch (Exception x)
      {
        Logger.error(x);
        return;
      }

      model.setDocumentModified(false);
      model.close();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + model.hashCode() + ", " + ext
        + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein Event, das das übergebene Dokument in eine temporäre Datei speichert
   * und eine externe Anwendung mit dieser aufruft, wobei der
   * ExterneAnwendungen-Abschnitt zu ext die näheren Details wie den FILTER regelt.
   * 
   * @param model
   *          Das an die externe Anwendung weiterzureichende TextDocumentModel.
   * @param ext
   *          identifiziert den entsprechenden Eintrag im Abschnitt
   *          ExterneAnwendungen.
   */
  public static void handleSaveTempAndOpenExt(TextDocumentModel model, String ext)
  {
    handle(new OnSaveTempAndOpenExt(model, ext));
  }

  /**
   * Dieses Event wird vom FormModelImpl ausgelöst, wenn der Benutzer die Aktion
   * "saveTempAndOpenExt" aktiviert hat.
   * 
   * @author matthias.benkmann
   */
  private static class OnSaveTempAndOpenExt extends BasicEvent
  {
    private TextDocumentModel model;

    private String ext;

    public OnSaveTempAndOpenExt(TextDocumentModel model, String ext)
    {
      this.model = model;
      this.ext = ext;
    }

    @Override
    protected void doit()
    {
      try
      {
        OpenExt openExt = new OpenExt(ext, WollMuxFiles.getWollmuxConf());
        openExt.setSource(UNO.XStorable(model.doc));
        openExt.storeIfNecessary();
        openExt.launch(new OpenExt.ExceptionHandler()
        {
          @Override
          public void handle(Exception x)
          {
            Logger.error(x);
          }
        });
      }
      catch (Exception x)
      {
        Logger.error(x);
        return;
      }
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + model.hashCode() + ", " + ext
        + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das ggf. notwendige interaktive
   * Initialisierungen vornimmt. Derzeit wird vor allem die Konsitenz der
   * persönlichen Absenderliste geprüft und der AbsenderAuswählen Dialog gestartet,
   * falls die Liste leer ist.
   */
  public static void handleInitialize()
  {
    handle(new OnInitialize());
  }

  /**
   * Dieses Event wird als erstes WollMuxEvent bei der Initialisierung des WollMux im
   * WollMuxSingleton erzeugt und übernimmt alle benutzersichtbaren (interaktiven)
   * Initialisierungen wie z.B. das Darstellen des AbsenderAuswählen-Dialogs, falls
   * die PAL leer ist.
   * 
   * @author christoph.lutz TESTED
   */
  private static class OnInitialize extends BasicEvent
  {
    @Override
    protected void doit()
    {
      DatasourceJoiner dsj = WollMuxFiles.getDatasourceJoiner();

      if (dsj.getLOS().size() == 0)
      {
        // falls es keine Datensätze im LOS gibt:
        // Die initialen Daten nach Heuristik versuchen zu finden:
        int found = searchDefaultSender(dsj);

        // Absender Auswählen Dialog starten:
        // wurde genau ein Datensatz gefunden, kann davon ausgegangen werden,
        // dass dieser OK ist - der Dialog muss dann nicht erscheinen.
        if (found != 1)
          WollMuxEventHandler.handleShowDialogAbsenderAuswaehlen();
        else
          handlePALChangedNotify();
      }
      else
      {
        // Liste der nicht zuordnenbaren Datensätze erstellen und ausgeben:
        String names = "";
        List<String> lost = WollMuxFiles.getLostDatasetDisplayStrings();
        if (lost.size() > 0)
        {
          for (String l : lost)
            names += "- " + l + "\n";
          String message =
            L.m("Die folgenden Datensätze konnten nicht "
              + "aus der Datenbank aktualisiert werden:\n\n" + "%1\n"
              + "Wenn dieses Problem nicht temporärer "
              + "Natur ist, sollten Sie diese Datensätze aus "
              + "ihrer Absenderliste löschen und neu hinzufügen!", names);
          ModalDialogs.showInfoModal(L.m("WollMux-Info"), message);
        }
      }
    }

    /**
     * Wertet den Konfigurationsabschnitt
     * PersoenlicheAbsenderlisteInitialisierung/Suchstrategie aus und versucht nach
     * der angegebenen Strategie (mindestens) einen Datensatz im DJ dsj zu finden,
     * der den aktuellen Benutzer repräsentiert. Fehlt der Konfigurationsabschnitt,
     * so wird die Defaultsuche BY_OOO_USER_PROFILE(Vorname "${givenname}" Nachname
     * "${sn}") gestartet. Liefert ein Element der Suchstrategie mindestens einen
     * Datensatz zurück, so werden die anderen Elemente der Suchstrategie nicht mehr
     * ausgewertet.
     * 
     * @param dsj
     *          Der DatasourceJoiner, in dem nach dem aktuellen Benutzer gesucht
     *          wird.
     * @return liefert die Anzahl der Datensätze, die nach Durchlaufen der
     *         Suchstrategie gefunden wurden.
     */
    private int searchDefaultSender(DatasourceJoiner dsj)
    {
      // Auswertung des Abschnitts
      // PersoenlicheAbsenderlisteInitialisierung/Suchstrategie
      ConfigThingy wmConf = WollMuxFiles.getWollmuxConf();
      ConfigThingy strat = null;
      try
      {
        strat =
          wmConf.query("PersoenlicheAbsenderlisteInitialisierung").query(
            "Suchstrategie").getLastChild();
      }
      catch (NodeNotFoundException e)
      {}

      if (strat != null)
      {
        // Suche über Suchstrategie aus Konfiguration
        for (Iterator<ConfigThingy> iter = strat.iterator(); iter.hasNext();)
        {
          ConfigThingy element = iter.next();
          int found = 0;
          if (element.getName().equals("BY_JAVA_PROPERTY"))
          {
            found = new ByJavaPropertyFinder(dsj).find(element);
          }
          else if (element.getName().equals("BY_OOO_USER_PROFILE"))
          {
            found = new ByOOoUserProfileFinder(dsj).find(element);
          }
          else
          {
            Logger.error(L.m("Ungültiger Schlüssel in Suchstrategie: %1",
              element.stringRepresentation()));
          }
          if (found != 0) return found;
        }
      }
      else
      {
        // Standardsuche über das OOoUserProfile:
        return new ByOOoUserProfileFinder(dsj).find("Vorname", "${givenname}",
          "Nachname", "${sn}");
      }

      return 0;
    }

    /**
     * Ein DataFinder sucht Datensätze im übergebenen dsj, wobei in der Beschreibung
     * der gesuchten Werte Variablen in der Form "${varname}" verwendet werden
     * können, die vor der Suche in einer anderen Datenquelle aufgelöst werden. Die
     * Auflösung erledigt durch die konkrete Klasse.
     */
    private static abstract class DataFinder
    {
      private DatasourceJoiner dsj;

      public DataFinder(DatasourceJoiner dsj)
      {
        this.dsj = dsj;
      }

      /**
       * Erwartet ein ConfigThingy, das ein oder zwei Schlüssel-/Wertpaare enthält
       * (in der Form "<KNOTEN>(<dbSpalte1> 'wert1' [<dbSpalte2> 'wert2'])" nach
       * denen in der Datenquelle gesucht werden soll. Die Beiden Wertpaare werden
       * dabei UND verknüpft. Die Werte wert1 und wert2 können über die Syntax
       * "${name}" Variablen referenzieren, die vor der Suche aufgelöst werden.
       * 
       * @param conf
       *          Das ConfigThingy, das die Suchabfrage beschreibt.
       * @return Die Anzahl der gefundenen Datensätze.
       */
      public int find(ConfigThingy conf)
      {
        int count = 0;
        String id1 = "";
        String id2 = "";
        String value1 = "";
        String value2 = "";
        for (Iterator<ConfigThingy> iter = conf.iterator(); iter.hasNext();)
        {
          ConfigThingy element = iter.next();
          if (count == 0)
          {
            id1 = element.getName();
            value1 = element.toString();
            count++;
          }
          else if (count == 1)
          {
            id2 = element.getName();
            value2 = element.toString();
            count++;
          }
          else
          {
            Logger.error(L.m("Nur max zwei Schlüssel/Wert-Paare werden als Argumente für Suchanfragen akzeptiert!"));
          }
        }

        if (count == 1)
        {
          return find(id1, value1);
        }
        else if (count == 2)
        {
          return find(id1, value1, id2, value2);
        }
        return 0;
      }

      /**
       * Sucht in der Datenquelle nach Datensätzen deren Feld dbSpalte den
       * evaluierten Wert von value enthält und überträgt die gefundenen Werte in die
       * PAL.
       * 
       * @param dbSpalte
       *          der Feldname über den nach dem evaluierten Wert von value gesucht
       *          wird.
       * @param value
       *          value wird vor der Suche mittels evaluate() evaluiert (d.h. evtl.
       *          vorhandene Variablen durch die entsprechenden Inhalte ersetzt
       *          ersetzt).
       * @return die Anzahl der gefundenen Datensätze
       */
      protected int find(String dbSpalte, String value)
      {
        Logger.debug2(this.getClass().getSimpleName() + ".tryToFind(" + dbSpalte
          + " '" + value + "')");
        try
        {
          String v = evaluate(value);
          if (v.length() == 0) return 0;
          QueryResults r = dsj.find(dbSpalte, v);
          return addToPAL(r);
        }
        catch (TimeoutException e)
        {
          Logger.error(e);
        }
        catch (IllegalArgumentException e)
        {
          Logger.debug(e);
        }
        return 0;
      }

      /**
       * Sucht in der Datenquelle nach Datensätzen wobei die beiden Suchbedingungen
       * (dbSpalte1==evaluate(value1) und dbSpalte2==evaluate(value2)) mit UND
       * verknüpft sind - die gefundenen Werte werden danach in die PAL kopiert.
       * 
       * @param dbSpalte1
       *          der Feldname über den nach dem evaluierten Wert von value gesucht
       *          wird.
       * @param value1
       *          value wird vor der Suche mittels evaluate() evaluiert (d.h. evtl.
       *          vorhandene Variablen durch die entsprechenden Inhalte ersetzt
       *          ersetzt).
       * @param dbSpalte2
       *          der Feldname über den nach dem evaluierten Wert von value gesucht
       *          wird.
       * @param value2
       *          value wird vor der Suche mittels evaluate() evaluiert (d.h. evtl.
       *          vorhandene Variablen durch die entsprechenden Inhalte ersetzt
       *          ersetzt).
       * @return die Anzahl der gefundenen Datensätze
       */
      protected int find(String dbSpalte1, String value1, String dbSpalte2,
          String value2)
      {
        Logger.debug2(this.getClass().getSimpleName() + ".tryToFind(" + dbSpalte1
          + " '" + value1 + "' " + dbSpalte2 + " '" + value2 + "')");
        try
        {
          String v1 = evaluate(value1);
          String v2 = evaluate(value2);
          if (v1.length() == 0 || v2.length() == 0) return 0;
          QueryResults r = dsj.find(dbSpalte1, v1, dbSpalte2, v2);
          return addToPAL(r);
        }
        catch (TimeoutException e)
        {
          Logger.error(e);
        }
        catch (IllegalArgumentException e)
        {
          Logger.debug(e);
        }
        return 0;
      }

      /**
       * Kopiert alle matches von QueryResults in die PAL.
       */
      private int addToPAL(QueryResults r)
      {
        for (Iterator<Dataset> iter = r.iterator(); iter.hasNext();)
        {
          DJDataset element = (DJDataset) iter.next();
          element.copy();
        }
        return r.size();
      }

      /**
       * Ersetzt die Variablen in exp durch deren evaluierten Inhalt, wobei die
       * Evaluierung über getValueForKey() erfolgt, die von jeder konkreten Klasse
       * implementiert wird. Evaluate() stellt auch sicher, dass die von
       * getValueForKey() zurückgelieferten Werte nicht selbst Variablen enthalten
       * können (indem die Variablenbegrenzer "${" und "}" durch "<" bzw. ">" ersetzt
       * werden.
       * 
       * @param exp
       *          der zu evaluierende Ausdruck
       * @return
       */
      protected String evaluate(String exp)
      {
        final Pattern VAR_PATTERN = Pattern.compile("\\$\\{([^\\}]*)\\}");
        while (true)
        {
          Matcher m = VAR_PATTERN.matcher(exp);
          if (!m.find()) break;
          String key = m.group(1);
          String value = getValueForKey(key);
          // keine Variablenbegrenzer "${" und "}" in value zulassen:
          value = value.replaceAll("\\$\\{", "<");
          value = value.replaceAll("\\}", ">");
          exp = m.replaceFirst(value);
        }
        return exp;
      }

      /**
       * Liefert den Wert zu einer Variable namens key und muss von jeder konkreten
       * Finder-Klasse implementiert werden.
       * 
       * @param key
       *          Der Schlüssel, zu dem der Wert zurückgeliefert werden soll.
       * @return der zugehörige Wert zum Schlüssel key.
       */
      protected abstract String getValueForKey(String key);
    }

    /**
     * Ein konkreter DataFinder, der für die Auflösung der Variable in getValueForKey
     * im Benutzerprofil der OOo Registry nachschaut (das selbe wie
     * Extras->Optionen->OpenOffice.org->Benutzerdaten).
     * 
     * @author christoph.lutz
     */
    private static class ByOOoUserProfileFinder extends DataFinder
    {
      public ByOOoUserProfileFinder(DatasourceJoiner dsj)
      {
        super(dsj);
      }

      @Override
      protected String getValueForKey(String key)
      {
        try
        {
          UnoService confProvider =
            UnoService.createWithContext(
              "com.sun.star.configuration.ConfigurationProvider",
              WollMuxSingleton.getInstance().getXComponentContext());

          UnoService confView =
            confProvider.create(
              "com.sun.star.configuration.ConfigurationAccess",
              new UnoProps("nodepath", "/org.openoffice.UserProfile/Data").getProps());
          return confView.xNameAccess().getByName(key).toString();
        }
        catch (Exception e)
        {
          Logger.error(
            L.m(
              "Konnte den Wert zum Schlüssel '%1' des OOoUserProfils nicht bestimmen:",
              key), e);
        }
        return "";
      }
    }

    /**
     * Ein konkreter DataFinder, der für die Auflösung der Variable in getValueForKey
     * die Methode System.getProperty(key) verwendet.
     * 
     * @author christoph.lutz
     */
    private static class ByJavaPropertyFinder extends DataFinder
    {
      public ByJavaPropertyFinder(DatasourceJoiner dsj)
      {
        super(dsj);
      }

      @Override
      protected String getValueForKey(String key)
      {
        try
        {
          return System.getProperty(key);
        }
        catch (java.lang.Exception e)
        {
          Logger.error(
            L.m("Konnte den Wert der JavaProperty '%1' nicht bestimmen:", key), e);
        }
        return "";
      }
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent zum Registrieren des übergebenen
   * XPALChangeEventListeners.
   * 
   * @param listener
   */
  public static void handleAddPALChangeEventListener(
      XPALChangeEventListener listener, Integer wollmuxConfHashCode)
  {
    handle(new OnAddPALChangeEventListener(listener, wollmuxConfHashCode));
  }

  /**
   * Dieses Event wird ausgelöst, wenn sich ein externer PALChangeEventListener beim
   * WollMux-Service registriert. Es sorgt dafür, dass der PALChangeEventListener in
   * die Liste der registrierten PALChangeEventListener im WollMuxSingleton
   * aufgenommen wird.
   * 
   * @author christoph.lutz
   */
  private static class OnAddPALChangeEventListener extends BasicEvent
  {
    private XPALChangeEventListener listener;

    private Integer wollmuxConfHashCode;

    public OnAddPALChangeEventListener(XPALChangeEventListener listener,
        Integer wollmuxConfHashCode)
    {
      this.listener = listener;
      this.wollmuxConfHashCode = wollmuxConfHashCode;
    }

    @Override
    protected void doit()
    {
      PersoenlicheAbsenderliste.getInstance().addPALChangeEventListener(listener);

      WollMuxEventHandler.handlePALChangedNotify();

      // Konsistenzprüfung: Stimmt WollMux-Konfiguration der entfernten
      // Komponente mit meiner Konfiguration überein? Ansonsten Fehlermeldung.
      if (wollmuxConfHashCode != null)
      {
        int myWmConfHash =
          WollMuxFiles.getWollmuxConf().stringRepresentation().hashCode();
        if (myWmConfHash != wollmuxConfHashCode.intValue())
          errorMessage(new InvalidBindingStateException(
            L.m("Die Konfiguration des WollMux muss neu eingelesen werden.\n\nBitte beenden Sie den WollMux und OpenOffice.org und schießen Sie alle laufenden 'soffice.bin'-Prozesse über den Taskmanager ab.")));
      }

    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + listener.hashCode() + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das den übergebenen XPALChangeEventListener
   * deregistriert.
   * 
   * @param listener
   *          der zu deregistrierende XPALChangeEventListener
   */
  public static void handleRemovePALChangeEventListener(
      XPALChangeEventListener listener)
  {
    handle(new OnRemovePALChangeEventListener(listener));
  }

  /**
   * Dieses Event wird vom WollMux-Service (...comp.WollMux) ausgelöst wenn sich ein
   * externe XPALChangeEventListener beim WollMux deregistriert. Der zu entfernende
   * XPALChangeEventListerner wird anschließend im WollMuxSingleton aus der Liste der
   * registrierten XPALChangeEventListener genommen.
   * 
   * @author christoph.lutz
   */
  private static class OnRemovePALChangeEventListener extends BasicEvent
  {
    private XPALChangeEventListener listener;

    public OnRemovePALChangeEventListener(XPALChangeEventListener listener)
    {
      this.listener = listener;
    }

    @Override
    protected void doit()
    {
      PersoenlicheAbsenderliste.getInstance().removePALChangeEventListener(listener);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + listener.hashCode() + ")";
    }
  }

  // *******************************************************************************************
  /**
   * Erzeugt ein neues WollMuxEvent zum Registrieren eines (frischen)
   * {@link DispatchProviderAndInterceptor} auf frame.
   * 
   * @param frame
   *          der {@link XFrame} auf den der {@link DispatchProviderAndInterceptor}
   *          registriert werden soll.
   */
  public static void handleRegisterDispatchInterceptor(XFrame frame)
  {
    if (frame == null)
      Logger.debug(L.m("Ignoriere handleRegisterDispatchInterceptor(null)"));
    else
      handle(new OnRegisterDispatchInterceptor(frame));
  }

  private static class OnRegisterDispatchInterceptor extends BasicEvent
  {
    private XFrame frame;

    public OnRegisterDispatchInterceptor(XFrame frame)
    {
      this.frame = frame;
    }

    @Override
    protected void doit()
    {
      try
      {
        DispatchProviderAndInterceptor.registerDocumentDispatchInterceptor(frame);
      }
      catch (java.lang.Exception e)
      {
        Logger.error(L.m("Kann DispatchInterceptor nicht registrieren:"), e);
      }

      // Sicherstellen, dass die Schaltflächen der Symbolleisten aktiviert werden:
      try
      {
        frame.contextChanged();
      }
      catch (java.lang.Exception e)
      {}
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + frame.hashCode() + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent zum Registrieren des übergebenen XEventListeners
   * und wird vom WollMux-Service aufgerufen.
   * 
   * @param listener
   *          der zu registrierende XEventListener.
   */
  public static void handleAddDocumentEventListener(XEventListener listener)
  {
    handle(new OnAddDocumentEventListener(listener));
  }

  private static class OnAddDocumentEventListener extends BasicEvent
  {
    private XEventListener listener;

    public OnAddDocumentEventListener(XEventListener listener)
    {
      this.listener = listener;
    }

    @Override
    protected void doit()
    {
      DocumentManager.getDocumentManager().addDocumentEventListener(listener);

      List<XComponent> processedDocuments = new Vector<XComponent>();
      DocumentManager.getDocumentManager().getProcessedDocuments(processedDocuments);

      for (XComponent compo : processedDocuments)
      {
        handleNotifyDocumentEventListener(listener, ON_WOLLMUX_PROCESSING_FINISHED,
          compo);
      }
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + listener.hashCode() + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das den übergebenen XEventListener zu
   * deregistriert.
   * 
   * @param listener
   *          der zu deregistrierende XEventListener
   */
  public static void handleRemoveDocumentEventListener(XEventListener listener)
  {
    handle(new OnRemoveDocumentEventListener(listener));
  }

  private static class OnRemoveDocumentEventListener extends BasicEvent
  {
    private XEventListener listener;

    public OnRemoveDocumentEventListener(XEventListener listener)
    {
      this.listener = listener;
    }

    @Override
    protected void doit()
    {
      DocumentManager.getDocumentManager().removeDocumentEventListener(listener);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + listener.hashCode() + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Über dieses Event werden alle registrierten DocumentEventListener (falls
   * listener==null) oder ein bestimmter registrierter DocumentEventListener (falls
   * listener != null) (XEventListener-Objekte) über Statusänderungen der
   * Dokumentbearbeitung informiert
   * 
   * @param listener
   *          der zu benachrichtigende XEventListener. Falls null werden alle
   *          registrierten Listener benachrichtig. listener wird auf jeden Fall nur
   *          benachrichtigt, wenn er zur Zeit der Abarbeitung des Events noch
   *          registriert ist.
   * @param eventName
   *          Name des Events
   * @param source
   *          Das von der Statusänderung betroffene Dokument (üblicherweise eine
   *          XComponent)
   * 
   * @author Christoph Lutz (D-III-ITD-5.1)
   */
  public static void handleNotifyDocumentEventListener(XEventListener listener,
      String eventName, Object source)
  {
    handle(new OnNotifyDocumentEventListener(listener, eventName, source));
  }

  private static class OnNotifyDocumentEventListener extends BasicEvent
  {
    private String eventName;

    private Object source;

    private XEventListener listener;

    public OnNotifyDocumentEventListener(XEventListener listener, String eventName,
        Object source)
    {
      this.listener = listener;
      this.eventName = eventName;
      this.source = source;
    }

    @Override
    protected void doit()
    {
      final com.sun.star.document.EventObject eventObject =
        new com.sun.star.document.EventObject();
      eventObject.Source = source;
      eventObject.EventName = eventName;

      Iterator<XEventListener> i =
          DocumentManager.getDocumentManager().documentEventListenerIterator();
      while (i.hasNext())
      {
        Logger.debug2("notifying XEventListener (event '" + eventName + "')");
        try
        {
          final XEventListener listener = i.next();
          if (this.listener == null || this.listener == listener) new Thread()
          {
            @Override
            public void run()
            {
              try
              {
                listener.notifyEvent(eventObject);
              }
              catch (java.lang.Exception x)
              {}
            }
          }.start();
        }
        catch (java.lang.Exception e)
        {
          i.remove();
        }
      }

      XComponent compo = UNO.XComponent(source);
      if (compo != null && eventName.equals(ON_WOLLMUX_PROCESSING_FINISHED))
        DocumentManager.getDocumentManager().setProcessingFinished(
          compo);

    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "('" + eventName + "', "
        + ((source != null) ? "#" + source.hashCode() : "null") + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent das signaisiert, dass die Druckfunktion
   * aufgerufen werden soll, die im TextDocumentModel model aktuell definiert ist.
   * Die Methode erwartet, dass vor dem Aufruf geprüft wurde, ob model eine
   * Druckfunktion definiert. Ist dennoch keine Druckfunktion definiert, so erscheint
   * eine Fehlermeldung im Log.
   * 
   * Das Event wird ausgelöst, wenn der registrierte WollMuxDispatchInterceptor eines
   * Dokuments eine entsprechende Nachricht bekommt.
   */
  public static void handleExecutePrintFunctions(TextDocumentModel model)
  {
    handle(new OnExecutePrintFunction(model));
  }

  private static class OnExecutePrintFunction extends BasicEvent
  {
    private TextDocumentModel model;

    public OnExecutePrintFunction(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      // Prüfen, ob alle gesetzten Druckfunktionen im aktuellen Kontext noch
      // Sinn machen:
      checkPrintPreconditions(model);
      stabilize();

      // Die im Dokument gesetzten Druckfunktionen ausführen:
      final XPrintModel pmod = model.createPrintModel(true);

      // Drucken im Hintergrund, damit der WollMuxEventHandler weiterläuft.
      new Thread()
      {
        @Override
        public void run()
        {
          pmod.printWithProps();
        }
      }.start();
    }

    /**
     * Es kann sein, dass zum Zeitpunkt des Drucken-Aufrufs eine Druckfunktion
     * gesetzt hat, die in der aktuellen Situation nicht mehr sinnvoll ist; Dieser
     * Umstand wird in checkPreconditons geprüft und die betroffene Druckfunktion
     * ggf. aus der Liste der Druckfunktionen entfernt.
     * 
     * @param printFunctions
     *          Menge der aktuell gesetzten Druckfunktionen.
     * 
     * @author Christoph Lutz (D-III-ITD-5.1)
     */
    protected static void checkPrintPreconditions(TextDocumentModel model)
    {
      Set<String> printFunctions = model.getPrintFunctions();

      // Ziffernanpassung der Sachleitenden Verfügungen durlaufen lassen, um zu
      // erkennen, ob Verfügungspunkte manuell aus dem Dokument gelöscht
      // wurden ohne die entsprechenden Knöpfe zum Einfügen/Entfernen von
      // Ziffern zu drücken.
      if (printFunctions.contains(SachleitendeVerfuegung.PRINT_FUNCTION_NAME))
      {
        SachleitendeVerfuegung.ziffernAnpassen(model);
      }

      // ...Platz für weitere Prüfungen.....
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Diese Methode erzeugt ein neues WollMuxEvent, mit dem die Eigenschaften der
   * Druckblöcke (z.B. allVersions) gesetzt werden können.
   * 
   * Das Event dient als Hilfe für die Komfortdruckfunktionen und wird vom
   * XPrintModel aufgerufen und mit diesem synchronisiert.
   * 
   * @param blockName
   *          Der Blocktyp dessen Druckblöcke behandelt werden sollen.
   * @param visible
   *          Der Block wird sichtbar, wenn visible==true und unsichtbar, wenn
   *          visible==false.
   * @param showHighlightColor
   *          gibt an ob die Hintergrundfarbe angezeigt werden soll (gilt nur, wenn
   *          zu einem betroffenen Druckblock auch eine Hintergrundfarbe angegeben
   *          ist).
   * 
   * @author Christoph Lutz (D-III-ITD-5.1)
   */
  public static void handleSetPrintBlocksPropsViaPrintModel(XTextDocument doc,
      String blockName, boolean visible, boolean showHighlightColor,
      ActionListener listener)
  {
    handle(new OnSetPrintBlocksPropsViaPrintModel(doc, blockName, visible,
      showHighlightColor, listener));
  }

  private static class OnSetPrintBlocksPropsViaPrintModel extends BasicEvent
  {
    private XTextDocument doc;

    private String blockName;

    private boolean visible;

    private boolean showHighlightColor;

    private ActionListener listener;

    public OnSetPrintBlocksPropsViaPrintModel(XTextDocument doc, String blockName,
        boolean visible, boolean showHighlightColor, ActionListener listener)
    {
      this.doc = doc;
      this.blockName = blockName;
      this.visible = visible;
      this.showHighlightColor = showHighlightColor;
      this.listener = listener;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      TextDocumentModel model =
        DocumentManager.getTextDocumentModel(doc);
      try
      {
        model.setPrintBlocksProps(blockName, visible, showHighlightColor);
      }
      catch (java.lang.Exception e)
      {
        errorMessage(e);
      }

      stabilize();
      if (listener != null) listener.actionPerformed(null);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + doc.hashCode() + ", '"
        + blockName + "', '" + visible + "', '" + showHighlightColor + "')";
    }
  }

  // *******************************************************************************************

  /**
   * Diese Methode erzeugt ein neues WollMux-Event über das die Liste der dem
   * Dokument doc zugeordneten Druckfunktionen verwaltet werden kann; ist
   * remove==false, so wird die Druckfunktion functionName in die Liste der
   * Druckfunktionen für dieses Dokument aufgenommen; ist remove==true, so wird die
   * Druckfunktion aus der Liste entfernt.
   * 
   * @param doc
   *          beschreibt das Dokument dessen Druckfunktionen verwaltet werden sollen.
   * @param functionName
   *          der Name der Druckfunktion, die hinzugefügt oder entfernt werden soll.
   * @param remove
   *          ist remove==false, so wird die Druckfunktion functionName in die Liste
   *          der Druckfunktionen für dieses Dokument aufgenommen; ist remove==true,
   *          so wird die Druckfunktion aus der Liste entfernt.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static void handleManagePrintFunction(XTextDocument doc,
      String functionName, boolean remove)
  {
    handle(new OnManagePrintFunction(doc, functionName, remove));
  }

  private static class OnManagePrintFunction extends BasicEvent
  {
    private XTextDocument doc;

    private String functionName;

    private boolean remove;

    public OnManagePrintFunction(XTextDocument doc, String functionName,
        boolean remove)
    {
      this.doc = doc;
      this.functionName = functionName;
      this.remove = remove;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      TextDocumentModel model =
        DocumentManager.getTextDocumentModel(doc);
      if (remove)
        model.removePrintFunction(functionName);
      else
        model.addPrintFunction(functionName);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + doc.hashCode() + ", '"
        + functionName + "', remove=" + remove + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das signasisiert, dass eine weitere Ziffer der
   * Sachleitenden Verfügungen eingefügt werden, bzw. eine bestehende Ziffer gelöscht
   * werden soll.
   * 
   * Das Event wird von WollMux.dispatch(...) geworfen, wenn Aufgrund eines Drucks
   * auf den Knopf der OOo-Symbolleiste ein "wollmux:ZifferEinfuegen" dispatch
   * erfolgte.
   */
  public static void handleButtonZifferEinfuegenPressed(TextDocumentModel model)
  {
    handle(new OnZifferEinfuegen(model));
  }

  private static class OnZifferEinfuegen extends BasicEvent
  {
    private TextDocumentModel model;

    public OnZifferEinfuegen(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      XTextCursor viewCursor = model.getViewCursor();
      if (viewCursor != null)
      {
        XTextRange vc =
          SachleitendeVerfuegung.insertVerfuegungspunkt(model, viewCursor);
        if (vc != null) viewCursor.gotoRange(vc, false);
      }

      stabilize();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das signasisiert, dass eine Abdruckzeile der
   * Sachleitenden Verfügungen eingefügt werden, bzw. eine bestehende Abdruckzeile
   * gelöscht werden soll.
   * 
   * Das Event wird von WollMux.dispatch(...) geworfen, wenn Aufgrund eines Drucks
   * auf den Knopf der OOo-Symbolleiste ein "wollmux:Abdruck" dispatch erfolgte.
   */
  public static void handleButtonAbdruckPressed(TextDocumentModel model)
  {
    handle(new OnAbdruck(model));
  }

  private static class OnAbdruck extends BasicEvent
  {
    private TextDocumentModel model;

    public OnAbdruck(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      XTextCursor viewCursor = model.getViewCursor();
      if (viewCursor != null)
      {
        XTextRange vc = SachleitendeVerfuegung.insertAbdruck(model, viewCursor);
        if (vc != null) viewCursor.gotoRange(vc, false);
      }

      stabilize();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das signasisiert, dass eine Zuleitungszeile der
   * Sachleitenden Verfügungen eingefügt werden, bzw. eine bestehende Zuleitungszeile
   * gelöscht werden soll.
   * 
   * Das Event wird von WollMux.dispatch(...) geworfen, wenn Aufgrund eines Drucks
   * auf den Knopf der OOo-Symbolleiste ein "wollmux:Zuleitungszeile" dispatch
   * erfolgte.
   */
  public static void handleButtonZuleitungszeilePressed(TextDocumentModel model)
  {
    handle(new OnButtonZuleitungszeilePressed(model));
  }

  private static class OnButtonZuleitungszeilePressed extends BasicEvent
  {
    private TextDocumentModel model;

    public OnButtonZuleitungszeilePressed(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      XTextCursor viewCursor = model.getViewCursor();
      if (viewCursor != null)
      {
        XTextRange vc =
          SachleitendeVerfuegung.insertZuleitungszeile(model, viewCursor);
        if (vc != null) viewCursor.gotoRange(vc, false);
      }

      stabilize();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das über den Bereich des viewCursors im Dokument
   * doc ein neues Bookmark mit dem Namen "WM(CMD'<blockname>')" legt, wenn nicht
   * bereits ein solches Bookmark im markierten Block definiert ist. Ist bereits ein
   * Bookmark mit diesem Namen vorhanden, so wird dieses gelöscht.
   * 
   * Das Event wird von WollMux.dispatch(...) geworfen, wenn Aufgrund eines Drucks
   * auf den Knopf der OOo-Symbolleiste ein "wollmux:markBlock#<blockname>" dispatch
   * erfolgte.
   * 
   * @param model
   *          Das Textdokument, in dem der Block eingefügt werden soll.
   * @param blockname
   *          Derzeit werden folgende Blocknamen akzeptiert "draftOnly",
   *          "notInOriginal", "originalOnly", "copyOnly" und "allVersions". Alle
   *          anderen Blocknamen werden ignoriert und keine Aktion ausgeführt.
   */
  public static void handleMarkBlock(TextDocumentModel model, String blockname)
  {
    handle(new OnMarkBlock(model, blockname));
  }

  private static class OnMarkBlock extends BasicEvent
  {
    private TextDocumentModel model;

    private String blockname;

    public OnMarkBlock(TextDocumentModel model, String blockname)
    {
      this.model = model;
      this.blockname = blockname;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      if (UNO.XBookmarksSupplier(model.doc) == null || blockname == null) return;

      ConfigThingy slvConf =
          WollMuxFiles.getWollmuxConf().query(
          "SachleitendeVerfuegungen");
      Integer highlightColor = null;

      XTextCursor range = model.getViewCursor();

      if (range == null) return;

      if (range.isCollapsed())
      {
        ModalDialogs.showInfoModal(L.m("Fehler"),
          L.m("Bitte wählen Sie einen Bereich aus, der markiert werden soll."));
        return;
      }

      String markChange = null;
      if (blockname.equalsIgnoreCase("allVersions"))
      {
        markChange = L.m("wird immer gedruckt");
        highlightColor = getHighlightColor(slvConf, "ALL_VERSIONS_HIGHLIGHT_COLOR");
      }
      else if (blockname.equalsIgnoreCase("draftOnly"))
      {
        markChange = L.m("wird nur im Entwurf gedruckt");
        highlightColor = getHighlightColor(slvConf, "DRAFT_ONLY_HIGHLIGHT_COLOR");
      }
      else if (blockname.equalsIgnoreCase("notInOriginal"))
      {
        markChange = L.m("wird im Original nicht gedruckt");
        highlightColor =
          getHighlightColor(slvConf, "NOT_IN_ORIGINAL_HIGHLIGHT_COLOR");
      }
      else if (blockname.equalsIgnoreCase("originalOnly"))
      {
        markChange = L.m("wird ausschließlich im Original gedruckt");
        highlightColor = getHighlightColor(slvConf, "ORIGINAL_ONLY_HIGHLIGHT_COLOR");
      }
      else if (blockname.equalsIgnoreCase("copyOnly"))
      {
        markChange = L.m("wird ausschließlich in Abdrucken gedruckt");
        highlightColor = getHighlightColor(slvConf, "COPY_ONLY_HIGHLIGHT_COLOR");
      }
      else
        return;

      String bookmarkStart = "WM(CMD '" + blockname + "'";
      String hcAtt = "";
      if (highlightColor != null)
      {
        String colStr = "00000000";
        colStr += Integer.toHexString(highlightColor.intValue());
        colStr = colStr.substring(colStr.length() - 8, colStr.length());
        hcAtt = " HIGHLIGHT_COLOR '" + colStr + "'";
      }
      String bookmarkName = bookmarkStart + hcAtt + ")";

      Pattern bookmarkPattern = DocumentCommands.getPatternForCommand(blockname);
      Set<String> bmNames =
        TextDocument.getBookmarkNamesMatching(bookmarkPattern, range);

      if (bmNames.size() > 0)
      {
        // bereits bestehende Blöcke löschen
        Iterator<String> iter = bmNames.iterator();
        while (iter.hasNext())
        {
          bookmarkName = iter.next();
          try
          {
            Bookmark b =
              new Bookmark(bookmarkName, UNO.XBookmarksSupplier(model.doc));
            if (bookmarkName.contains("HIGHLIGHT_COLOR"))
              UNO.setPropertyToDefault(b.getTextCursor(), "CharBackColor");
            b.remove();
          }
          catch (NoSuchElementException e)
          {}
        }
        ModalDialogs.showInfoModal(
          L.m("Markierung des Blockes aufgehoben"),
          L.m(
            "Der ausgewählte Block enthielt bereits eine Markierung 'Block %1'. Die bestehende Markierung wurde aufgehoben.",
            markChange));
      }
      else
      {
        // neuen Block anlegen
        model.addNewDocumentCommand(range, bookmarkName);
        if (highlightColor != null)
        {
          UNO.setProperty(range, "CharBackColor", highlightColor);
          // ViewCursor kollabieren, da die Markierung die Farben verfälscht
          // darstellt.
          XTextCursor vc = model.getViewCursor();
          if (vc != null) vc.collapseToEnd();
        }
        ModalDialogs.showInfoModal(L.m("Block wurde markiert"),
          L.m("Der ausgewählte Block %1.", markChange));
      }

      // PrintBlöcke neu einlesen:
      model.getDocumentCommands().update();
      DocumentCommandInterpreter dci =
        new DocumentCommandInterpreter(model, WollMuxFiles.isDebugMode());
      dci.scanGlobalDocumentCommands();
      dci.scanInsertFormValueCommands();

      stabilize();
    }

    /**
     * Liefert einen Integer der Form AARRGGBB (hex), der den Farbwert repräsentiert,
     * der in slvConf im Attribut attribute hinterlegt ist oder null, wenn das
     * Attribut nicht existiert oder der dort enthaltene String-Wert sich nicht in
     * eine Integerzahl konvertieren lässt.
     * 
     * @param slvConf
     * @param attribute
     */
    private static Integer getHighlightColor(ConfigThingy slvConf, String attribute)
    {
      try
      {
        String highlightColor = slvConf.query(attribute).getLastChild().toString();
        if (highlightColor.equals("") || highlightColor.equalsIgnoreCase("none"))
          return null;
        int hc = Integer.parseInt(highlightColor, 16);
        return Integer.valueOf(hc);
      }
      catch (NodeNotFoundException e)
      {
        return null;
      }
      catch (NumberFormatException e)
      {
        Logger.error(L.m("Der angegebene Farbwert im Attribut '%1' ist ungültig!",
          attribute));
        return null;
      }
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + model.hashCode() + ", '"
        + blockname + "')";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das dafür sorgt, dass eine Datei wollmux.dump
   * erzeugt wird, die viele für die Fehlersuche relevanten Informationen enthält wie
   * z.B. Versionsinfo, Inhalt der wollmux.conf, cache.conf, StringRepräsentation der
   * Konfiguration im Speicher und eine Kopie der Log-Datei.
   * 
   * Das Event wird von der WollMuxBar geworfen, die (speziell für Admins, nicht für
   * Endbenutzer) einen entsprechenden Button besitzt.
   */
  public static void handleDumpInfo()
  {
    handle(new OnDumpInfo());
  }

  private static class OnDumpInfo extends BasicEvent
  {

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      final String title = L.m("Fehlerinfos erstellen");

      String name = WollMuxFiles.dumpInfo();

      if (name != null)
        ModalDialogs.showInfoModal(
          title,
          L.m(
            "Die Fehlerinformationen des WollMux wurden erfolgreich in die Datei '%1' geschrieben.",
            name));
      else
        ModalDialogs.showInfoModal(
          title,
          L.m("Die Fehlerinformationen des WollMux konnten nicht geschrieben werden\n\nDetails siehe Datei wollmux.log!"));
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "()";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das signasisiert, dass das gesamte Office (und
   * damit auch der WollMux) OHNE Sicherheitsabfragen(!) beendet werden soll.
   * 
   * Das Event wird von der WollMuxBar geworfen, die (speziell für Admins, nicht für
   * Endbenutzer) einen entsprechenden Button besitzt.
   */
  public static void handleKill()
  {
    handle(new OnKill());
  }

  private static class OnKill extends BasicEvent
  {
    @Override
    protected void doit() throws WollMuxFehlerException
    {
      if (UNO.desktop != null)
      {
        UNO.desktop.terminate();
      }
      else
      {
        System.exit(0);
      }
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "()";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das dafür sorgt, dass im Textdokument doc das
   * Formularfeld mit der ID id auf den Wert value gesetzt wird. Ist das Dokument ein
   * Formulardokument (also mit einer angezeigten FormGUI), so wird die Änderung über
   * die FormGUI vorgenommen, die zugleich dafür sorgt, dass von id abhängige
   * Formularfelder mit angepasst werden. Besitzt das Dokument keine
   * Formularbeschreibung, so wird der Wert direkt gesetzt, ohne Äbhängigkeiten zu
   * beachten. Nach der erfolgreichen Ausführung aller notwendigen Anpassungen wird
   * der unlockActionListener benachrichtigt.
   * 
   * Das Event wird aus den Implementierungen von XPrintModel (siehe
   * TextDocumentModel) und XWollMuxDocument (siehe compo.WollMux) geworfen, wenn
   * dort die Methode setFormValue aufgerufen wird.
   * 
   * @param doc
   *          Das Dokument, in dem das Formularfeld mit der ID id neu gesetzt werden
   *          soll.
   * @param id
   *          Die ID des Formularfeldes, dessen Wert verändert werden soll. Ist die
   *          FormGUI aktiv, so werden auch alle von id abhängigen Formularwerte neu
   *          gesetzt.
   * @param value
   *          Der neue Wert des Formularfeldes id
   * @param unlockActionListener
   *          Der unlockActionListener wird immer informiert, wenn alle notwendigen
   *          Anpassungen durchgeführt wurden.
   */
  public static void handleSetFormValue(XTextDocument doc, String id, String value,
      ActionListener unlockActionListener)
  {
    handle(new OnSetFormValue(doc, id, value, unlockActionListener));
  }

  private static class OnSetFormValue extends BasicEvent
  {
    private XTextDocument doc;

    private String id;

    private String value;

    private final ActionListener listener;

    public OnSetFormValue(XTextDocument doc, String id, String value,
        ActionListener listener)
    {
      this.doc = doc;
      this.id = id;
      this.value = value;
      this.listener = listener;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      TextDocumentModel model =
        DocumentManager.getTextDocumentModel(doc);

      FormModel formModel = model.getFormModel();
      if (formModel != null)
      {
        // Werte über den FormController (den das FormModel kennt) setzen lassen
        // (damit sind auch automatisch alle Abhängigkeiten richtig aufgelöst)
        formModel.setValue(id, value, new ActionListener()
        {
          @Override
          public void actionPerformed(ActionEvent arg0)
          {
            handleSetFormValueFinished(listener);
          }
        });
      }
      else
      {
        // Werte selber setzen:
        model.setFormFieldValue(id, value);
        model.updateFormFields(id);
        if (listener != null) listener.actionPerformed(null);
      }
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + doc.hashCode() + ", id='" + id
        + "', value='" + value + "')";
    }
  }

  // *******************************************************************************************

  /**
   * Sammelt alle Formularfelder des Dokuments model auf, die nicht von
   * WollMux-Kommandos umgeben sind, jedoch trotzdem vom WollMux verstanden und
   * befüllt werden (derzeit c,s,s,t,textfield,Database-Felder).
   * 
   * Das Event wird aus der Implementierung von XPrintModel (siehe TextDocumentModel)
   * geworfen, wenn dort die Methode collectNonWollMuxFormFields aufgerufen wird.
   * 
   * @param model
   * @param unlockActionListener
   *          Der unlockActionListener wird immer informiert, wenn alle notwendigen
   *          Anpassungen durchgeführt wurden.
   */
  public static void handleCollectNonWollMuxFormFieldsViaPrintModel(
      TextDocumentModel model, ActionListener listener)
  {
    handle(new OnCollectNonWollMuxFormFieldsViaPrintModel(model, listener));
  }

  private static class OnCollectNonWollMuxFormFieldsViaPrintModel extends BasicEvent
  {
    private TextDocumentModel model;

    private ActionListener listener;

    public OnCollectNonWollMuxFormFieldsViaPrintModel(TextDocumentModel model,
        ActionListener listener)
    {
      this.model = model;
      this.listener = listener;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      model.collectNonWollMuxFormFields();

      stabilize();
      if (listener != null) listener.actionPerformed(null);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Dieses WollMuxEvent ist das Gegenstück zu handleSetFormValue und wird dann
   * erzeugt, wenn nach einer Änderung eines Formularwertes - gesteuert durch die
   * FormGUI - alle abhängigen Formularwerte angepasst wurden. In diesem Fall ist die
   * einzige Aufgabe dieses Events, den unlockActionListener zu informieren, den
   * handleSetFormValueViaPrintModel() nicht selbst informieren konnte.
   * 
   * Das Event wird aus der Implementierung vom OnSetFormValueViaPrintModel.doit()
   * erzeugt, wenn Feldänderungen über die FormGUI laufen.
   * 
   * @param unlockActionListener
   *          Der zu informierende unlockActionListener.
   */
  public static void handleSetFormValueFinished(ActionListener unlockActionListener)
  {
    handle(new OnSetFormValueFinished(unlockActionListener));
  }

  private static class OnSetFormValueFinished extends BasicEvent
  {
    private ActionListener listener;

    public OnSetFormValueFinished(ActionListener unlockActionListener)
    {
      this.listener = unlockActionListener;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      if (listener != null) listener.actionPerformed(null);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "()";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das dafür sorgt, dass im Textdokument doc alle
   * insertValue-Befehle mit einer DB_SPALTE, die in der übergebenen
   * mapDbSpalteToValue enthalten sind neu für den entsprechenden Wert evaluiert und
   * gesetzt werden, unabhängig davon, ob sie den Status DONE besitzen oder nicht.
   * 
   * Das Event wird aus der Implementierung von XWollMuxDocument (siehe
   * compo.WollMux) geworfen, wenn dort die Methode setInsertValue aufgerufen wird.
   * 
   * @param doc
   *          Das Dokument, in dem das die insertValue-Kommandos neu gesetzt werden
   *          sollen.
   * @param mapDbSpalteToValue
   *          Enthält eine Zuordnung von DB_SPALTEn auf die zu setzenden Werte.
   *          Enthält ein betroffenes Dokumentkommando eine Trafo, so wird die Trafo
   *          mit dem zugehörigen Wert ausgeführt und das Transformationsergebnis als
   *          neuer Inhalt des Bookmarks gesetzt.
   * @param unlockActionListener
   *          Wird informiert, sobald das Event vollständig abgearbeitet wurde.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static void handleSetInsertValues(XTextDocument doc,
      Map<String, String> mapDbSpalteToValue, ActionListener unlockActionListener)
  {
    handle(new OnSetInsertValues(doc, mapDbSpalteToValue, unlockActionListener));
  }

  private static class OnSetInsertValues extends BasicEvent
  {
    private XTextDocument doc;

    private Map<String, String> mapDbSpalteToValue;

    private ActionListener listener;

    public OnSetInsertValues(XTextDocument doc,
        Map<String, String> mapDbSpalteToValue, ActionListener unlockActionListener)
    {
      this.doc = doc;
      this.mapDbSpalteToValue = mapDbSpalteToValue;
      this.listener = unlockActionListener;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      TextDocumentModel model =
        DocumentManager.getTextDocumentModel(doc);

      for (DocumentCommand cmd : model.getDocumentCommands())
        // stellt sicher, dass listener am Schluss informiert wird
        try
        {
          if (cmd instanceof DocumentCommand.InsertValue)
          {
            DocumentCommand.InsertValue insVal = (DocumentCommand.InsertValue) cmd;
            String value = mapDbSpalteToValue.get(insVal.getDBSpalte());
            if (value != null)
            {
              value = model.getTransformedValue(insVal.getTrafoName(), value);
              if ("".equals(value))
              {
                cmd.setTextRangeString("");
              }
              else
              {
                cmd.setTextRangeString(insVal.getLeftSeparator() + value
                  + insVal.getRightSeparator());
              }
            }
          }
        }
        catch (java.lang.Exception e)
        {
          Logger.error(e);
        }
      if (listener != null) listener.actionPerformed(null);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + doc.hashCode()
        + ", Nr.Values=" + mapDbSpalteToValue.size() + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das signasisiert, das ein Textbaustein über den
   * Textbaustein-Bezeichner direkt ins Dokument eingefügt wird. Mit reprocess wird
   * übergeben, wann die Dokumentenkommandos ausgewertet werden soll. Mir reprocess =
   * true sofort.
   * 
   * Das Event wird von WollMux.dispatch(...) geworfen z.B über Druck eines
   * Tastenkuerzels oder Druck auf den Knopf der OOo-Symbolleiste ein
   * "wollmux:TextbausteinEinfuegen" dispatch erfolgte.
   */
  public static void handleTextbausteinEinfuegen(TextDocumentModel model,
      boolean reprocess)
  {
    handle(new OnTextbausteinEinfuegen(model, reprocess));
  }

  private static class OnTextbausteinEinfuegen extends BasicEvent
  {
    private TextDocumentModel model;

    private boolean reprocess;

    public OnTextbausteinEinfuegen(TextDocumentModel model, boolean reprocess)
    {
      this.model = model;
      this.reprocess = reprocess;

    }

    @Override
    protected void doit()
    {
      XTextCursor viewCursor = model.getViewCursor();
      try
      {
        TextModule.createInsertFragFromIdentifier(model.doc, viewCursor, reprocess);
        if (reprocess) handleReprocessTextDocument(model);
        if (!reprocess)
          ModalDialogs.showInfoModal(L.m("WollMux"),
            L.m("Der Textbausteinverweis wurde eingefügt."));
      }
      catch (WollMuxFehlerException e)
      {
        ModalDialogs.showInfoModal(L.m("WollMux-Fehler"), e.getMessage());
      }
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ", " + reprocess + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das signasisiert, das der nächste Platzhalter
   * ausgehend vom Cursor angesprungen wird * Das Event wird von
   * WollMux.dispatch(...) geworfen z.B über Druck eines Tastenkuerzels oder Druck
   * auf den Knopf der OOo-Symbolleiste ein "wollmux:PlatzhalterAnspringen" dispatch
   * erfolgte.
   */
  public static void handleJumpToPlaceholder(TextDocumentModel model)
  {
    handle(new OnJumpToPlaceholder(model));
  }

  private static class OnJumpToPlaceholder extends BasicEvent
  {
    private TextDocumentModel model;

    public OnJumpToPlaceholder(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      XTextCursor viewCursor = model.getViewCursor();

      try
      {
        TextModule.jumpPlaceholders(model.doc, viewCursor);
      }
      catch (java.lang.Exception e)
      {
        Logger.error(e);
      }

      stabilize();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das signasisiert, das die nächste Marke
   * 'setJumpMark' angesprungen werden soll. Wird im
   * DocumentCommandInterpreter.DocumentExpander.fillPlaceholders aufgerufen wenn
   * nach dem Einfügen von Textbausteine keine Einfügestelle vorhanden ist aber eine
   * Marke 'setJumpMark'
   */
  public static void handleJumpToMark(XTextDocument doc, boolean msg)
  {
    handle(new OnJumpToMark(doc, msg));
  }

  private static class OnJumpToMark extends BasicEvent
  {
    private XTextDocument doc;

    private boolean msg;

    public OnJumpToMark(XTextDocument doc, boolean msg)
    {
      this.doc = doc;
      this.msg = msg;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {

      TextDocumentModel model =
        DocumentManager.getTextDocumentModel(doc);

      XTextCursor viewCursor = model.getViewCursor();
      if (viewCursor == null) return;

      DocumentCommand cmd = model.getFirstJumpMark();

      if (cmd != null)
      {
        try
        {
          XTextRange range = cmd.getTextCursor();
          if (range != null) viewCursor.gotoRange(range.getStart(), false);
        }
        catch (java.lang.Exception e)
        {
          Logger.error(e);
        }

        boolean modified = model.getDocumentModified();
        cmd.markDone(true);
        model.setDocumentModified(modified);

        model.getDocumentCommands().update();

      }
      else
      {
        if (msg)
        {
          ModalDialogs.showInfoModal(L.m("WollMux"),
            L.m("Kein Platzhalter und keine Marke 'setJumpMark' vorhanden!"));
        }
      }

      stabilize();
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(#" + doc.hashCode() + ", " + msg
        + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das signasisiert, dass die neue
   * Seriendruckfunktion des WollMux gestartet werden soll.
   * 
   * Das Event wird über den DispatchHandler aufgerufen, wenn z.B. über das Menü
   * "Extras->Seriendruck (WollMux)" die dispatch-url wollmux:SeriendruckNeu
   * abgesetzt wurde.
   */
  public static void handleSeriendruck(TextDocumentModel model,
      boolean useDocPrintFunctions)
  {
    handle(new OnSeriendruck(model, useDocPrintFunctions));
  }

  private static class OnSeriendruck extends BasicEvent
  {
    private TextDocumentModel model;

    public OnSeriendruck(TextDocumentModel model, boolean useDocumentPrintFunctions)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      // Bestehenden Max in den Vordergrund holen oder neuen Max erzeugen.
      MailMergeNew mmn = model.getCurrentMailMergeNew();
      if (mmn != null)
      {
        return;
      }
      else
      {
        mmn = new MailMergeNew(model, new ActionListener()
        {
          @Override
          public void actionPerformed(ActionEvent actionEvent)
          {
            if (actionEvent.getSource() instanceof MailMergeNew)
              WollMuxEventHandler.handleMailMergeNewReturned(model);
          }
        });
        model.setCurrentMailMergeNew(mmn);
      }
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Der Handler für das Drucken eines TextDokuments führt in Abhängigkeit von der
   * Existenz von Serienbrieffeldern und Druckfunktion die entsprechenden Aktionen
   * aus.
   * 
   * Das Event wird über den DispatchHandler aufgerufen, wenn z.B. über das Menü
   * "Datei->Drucken" oder über die Symbolleiste die dispatch-url .uno:Print bzw.
   * .uno:PrintDefault abgesetzt wurde.
   */
  public static void handlePrint(TextDocumentModel model, XDispatch origDisp,
      com.sun.star.util.URL origUrl, PropertyValue[] origArgs)
  {
    handle(new OnPrint(model, origDisp, origUrl, origArgs));
  }

  private static class OnPrint extends BasicEvent
  {
    private TextDocumentModel model;

    private XDispatch origDisp;

    private com.sun.star.util.URL origUrl;

    private PropertyValue[] origArgs;

    public OnPrint(TextDocumentModel model, XDispatch origDisp,
        com.sun.star.util.URL origUrl, PropertyValue[] origArgs)
    {
      this.model = model;
      this.origDisp = origDisp;
      this.origUrl = origUrl;
      this.origArgs = origArgs;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      boolean hasPrintFunction = model.getPrintFunctions().size() > 0;

      if (Workarounds.applyWorkaroundForOOoIssue96281())
      {
        try
        {
          Object viewSettings =
            UNO.XViewSettingsSupplier(model.doc.getCurrentController()).getViewSettings();
          UNO.setProperty(viewSettings, "ZoomType", DocumentZoomType.BY_VALUE);
          UNO.setProperty(viewSettings, "ZoomValue", Short.valueOf((short) 100));
        }
        catch (java.lang.Exception e)
        {}
      }

      if (hasPrintFunction)
      {
        // Druckfunktion aufrufen
        handleExecutePrintFunctions(model);
      }
      else
      {
        // Forward auf Standardfunktion
        if (origDisp != null) origDisp.dispatch(origUrl, origArgs);
      }
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Der Handler für einen abgespeckten Speichern-Unter-Dialog des WollMux, der in
   * Abängigkeit von einer gesetzten FilenameGeneratorFunction über den WollMux
   * aufgrufen und mit dem generierten Filenamen vorbelegt wird.
   * 
   * Das Event wird über den DispatchHandler aufgerufen, wenn z.B. über das Menü
   * "Datei->SaveAs" oder über die Symbolleiste die dispatch-url .uno:Save bzw.
   * .uno:SaveAs abgesetzt wurde.
   */
  public static void handleSaveAs(TextDocumentModel model, XDispatch origDisp,
      com.sun.star.util.URL origUrl, PropertyValue[] origArgs)
  {
    handle(new OnSaveAs(model, origDisp, origUrl, origArgs));
  }

  private static class OnSaveAs extends BasicEvent
  {
    private TextDocumentModel model;

    private XDispatch origDisp;

    private com.sun.star.util.URL origUrl;

    private PropertyValue[] origArgs;

    public OnSaveAs(TextDocumentModel model, XDispatch origDisp,
        com.sun.star.util.URL origUrl, PropertyValue[] origArgs)
    {
      this.model = model;
      this.origDisp = origDisp;
      this.origUrl = origUrl;
      this.origArgs = origArgs;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      // FilenameGeneratorFunction auslesen und parsen
      FunctionLibrary lib = model.getFunctionLibrary();
      ConfigThingy funcConf = model.getFilenameGeneratorFunc();
      Function func = null;
      if (funcConf != null) try
      {
        func = FunctionFactory.parse(funcConf, lib, null, null);
      }
      catch (ConfigurationErrorException e)
      {
        Logger.error(L.m("Kann FilenameGeneratorFunction nicht parsen!"), e);
      }

      // Original-Dispatch ausführen, wenn keine FilenameGeneratorFunction gesetzt
      if (func == null)
      {
        if (origDisp != null) origDisp.dispatch(origUrl, origArgs);
        return;
      }

      boolean done = false;
      File file = ensureFileHasODTSuffix(getDefaultFile(func));
      JFileChooser fc = createODTFileChooser(file);
      while (!done)
      {
        done = true;
        if (fc.showSaveDialog(null) == JFileChooser.APPROVE_OPTION)
        {
          boolean save = true;
          File f = ensureFileHasODTSuffix(fc.getSelectedFile());

          // Sicherheitsabfage vor Überschreiben
          if (f.exists())
          {
            save = false;
            int res =
              JOptionPane.showConfirmDialog(
                null,
                L.m("Datei %1 existiert bereits. Soll sie überschrieben werden?",
                  f.getName()), L.m("Überschreiben?"),
                JOptionPane.YES_NO_CANCEL_OPTION);
            if (res == JOptionPane.NO_OPTION) done = false;
            if (res == JOptionPane.OK_OPTION) save = true;
          }

          if (save) saveAs(f);
        }
      }
    }

    private File getDefaultFile(Function func)
    {
      Map<String, String> fields = model.getFormFieldValues();
      Values.SimpleMap values = new Values.SimpleMap();
      for (String par : func.parameters())
      {
        String value = fields.get(par);
        if (value == null) value = "";
        values.put(par, value);
      }
      String filename = func.getString(values);
      File f = new File(filename);
      if (f.isAbsolute()) return f;

      try
      {
        Object ps = UNO.createUNOService("com.sun.star.util.PathSettings");
        URL dir = new URL(AnyConverter.toString(UNO.getProperty(ps, "Work")));
        return new File(dir.getPath(), filename);
      }
      catch (com.sun.star.lang.IllegalArgumentException e)
      {
        Logger.error(e);
      }
      catch (MalformedURLException e)
      {
        Logger.error(e);
      }
      return new File(filename);
    }

    private JFileChooser createODTFileChooser(File file)
    {
      JFileChooser fc = new JFileChooser()
      {
        private static final long serialVersionUID = 1560806929064954454L;

        // Laut Didi kommt der JFileChooser unter Windows nicht im Vordergrund.
        // Deshalb das Überschreiben der createDialog-Methode und Setzen von
        // alwaysOnTop(true)
        @Override
        protected JDialog createDialog(Component parent) throws HeadlessException
        {
          JDialog dialog = super.createDialog(parent);
          dialog.setAlwaysOnTop(true);
          return dialog;
        }
      };
      fc.setMultiSelectionEnabled(false);
      fc.setFileFilter(new FileFilter()
      {
        @Override
        public String getDescription()
        {
          return L.m("ODF Textdokument");
        }

        @Override
        public boolean accept(File f)
        {
          return f.getName().toLowerCase().endsWith(".odt") || f.isDirectory();
        }
      });
      fc.setSelectedFile(file);
      return fc;
    }

    private void saveAs(File f)
    {
      model.flushPersistentData();
      try
      {
        String url = UNO.getParsedUNOUrl(f.toURI().toURL().toString()).Complete;
        if (UNO.XStorable(model.doc) != null)
          UNO.XStorable(model.doc).storeAsURL(url, new PropertyValue[] {});
      }
      catch (MalformedURLException e)
      {
        Logger.error(L.m("das darf nicht passieren"), e);
      }
      catch (com.sun.star.io.IOException e)
      {
        Logger.error(e);
        JOptionPane.showMessageDialog(null, L.m(
          "Das Speichern der Datei %1 ist fehlgeschlagen!\n\n%2", f.toString(),
          e.getLocalizedMessage()), L.m("Fehler beim Speichern"),
          JOptionPane.ERROR_MESSAGE);
      }
    }

    private static File ensureFileHasODTSuffix(File f)
    {
      if (f != null && !f.getName().toLowerCase().endsWith(".odt"))
      {
        return new File(f.getParent(), f.getName() + ".odt");
      }
      return f;
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMuxEvent, das signasisiert, dass der FormController (der
   * zeitgleich mit einer FormGUI zum TextDocument model gestartet wird) vollständig
   * initialisiert ist und notwendige Aktionen wie z.B. das Zurücksetzen des
   * Modified-Status des Dokuments durchgeführt werden können. Vor dem Zurücksetzen
   * des Modified-Status, wird auf die erste Seite des Dokuments gesprungen.
   * 
   * Das Event wird vom FormModel erzeugt, wenn es vom FormController eine
   * entsprechende Nachricht erhält.
   */
  public static void handleFormControllerInitCompleted(TextDocumentModel model)
  {
    handle(new OnFormControllerInitCompleted(model));
  }

  private static class OnFormControllerInitCompleted extends BasicEvent
  {
    private TextDocumentModel model;

    public OnFormControllerInitCompleted(TextDocumentModel model)
    {
      this.model = model;
    }

    @Override
    protected void doit() throws WollMuxFehlerException
    {
      // Springt zum Dokumentenanfang
      try
      {
        model.getViewCursor().gotoRange(model.doc.getText().getStart(), false);
      }
      catch (java.lang.Exception e)
      {
        Logger.debug(e);
      }

      // Beim Öffnen eines Formulars werden viele Änderungen am Dokument
      // vorgenommen (z.B. das Setzen vieler Formularwerte), ohne dass jedoch
      // eine entsprechende Benutzerinteraktion stattgefunden hat. Der
      // Modified-Status des Dokuments wird daher zurückgesetzt, damit nur
      // wirkliche Interaktionen durch den Benutzer modified=true setzen.
      model.setDocumentModified(false);
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "(" + model + ")";
    }
  }

  // *******************************************************************************************

  /**
   * Erzeugt ein neues WollMux-Event, in dem geprüft wird, ob der WollMux korrekt
   * installiert ist und keine Doppel- oder Halbinstallationen vorliegen. Ist der
   * WollMux fehlerhaft installiert, erscheint eine Fehlermeldung mit entsprechenden
   * Hinweisen.
   * 
   * Das Event wird geworfen, wenn der WollMux startet.
   */
  public static void handleCheckInstallation()
  {
    handle(new OnCheckInstallation());
  }

  private static class OnCheckInstallation extends BasicEvent
  {
    @Override
    protected void doit() throws WollMuxFehlerException
    {
      // Standardwerte für den Warndialog:
      boolean showdialog = true;
      String title = L.m("Mehrfachinstallation des WollMux");
      String msg =
        L.m("Es wurden eine systemweite und eine benutzerlokale Installation des WollMux\n(oder Überreste von einer unvollständigen Deinstallation) gefunden.\nDiese Konstellation kann obskure Fehler verursachen.\n\nEntfernen Sie eine der beiden Installationen.\n\nDie wollmux.log enthält nähere Informationen zu den betroffenen Pfaden.");
      String logMsg = msg;

      // Abschnitt Dialoge/MehrfachinstallationWarndialog auswerten
      try
      {
        ConfigThingy warndialog =
          WollMuxFiles.getWollmuxConf().query("Dialoge").query(
            "MehrfachinstallationWarndialog").getLastChild();
        try
        {
          msg = warndialog.get("MSG").toString();
        }
        catch (NodeNotFoundException e)
        {
          showdialog = false;
        }
        try
        {
          title = warndialog.get("TITLE").toString();
        }
        catch (NodeNotFoundException e)
        {}
      }
      catch (NodeNotFoundException e)
      {
        // Ist der Abschnitt nicht vorhanden, so greifen Standardwerte.
      }

      // Infos der Installationen einlesen.
      List<WollMuxInstallationDescriptor> wmInsts = getInstallations();

      // Variablen recentInstPath / recentInstLastModified / shared / local bestimmen
      String recentInstPath = "";
      Date recentInstLastModified = null;
      boolean shared = false;
      boolean local = false;
      for (WollMuxInstallationDescriptor desc : wmInsts)
      {
        shared = shared || desc.isShared;
        local = local || !desc.isShared;
        if (recentInstLastModified == null
          || desc.date.compareTo(recentInstLastModified) > 0)
        {
          recentInstLastModified = desc.date;
          recentInstPath = desc.path;
        }
      }

      // Variable wrongInstList bestimmen:
      String otherInstsList = "";
      for (WollMuxInstallationDescriptor desc : wmInsts)
      {
        if (!desc.path.equals(recentInstPath))
          otherInstsList += "- " + desc.path + "\n";
      }

      // Im Fehlerfall Dialog und Fehlermeldung bringen.
      if (local && shared)
      {

        // Variablen in msg evaluieren:
        DateFormat f = DateFormat.getDateInstance();
        msg = msg.replaceAll("\\$\\{RECENT_INST_PATH\\}", recentInstPath);
        msg =
          msg.replaceAll("\\$\\{RECENT_INST_LAST_MODIFIED\\}",
            f.format(recentInstLastModified));
        msg = msg.replaceAll("\\$\\{OTHER_INSTS_LIST\\}", otherInstsList);

        logMsg +=
          "\n" + L.m("Die juengste WollMux-Installation liegt unter:") + "\n- "
            + recentInstPath + "\n"
            + L.m("Ausserdem wurden folgende WollMux-Installationen gefunden:")
            + "\n" + otherInstsList;
        Logger.error(logMsg);

        if (showdialog) ModalDialogs.showInfoModal(title, msg, 0);
      }
    }

    private static class WollMuxInstallationDescriptor
    {
      public String path;

      public Date date;

      public boolean isShared;

      public WollMuxInstallationDescriptor(String path, Date date, boolean isShared)
      {
        this.path = path;
        this.date = date;
        this.isShared = isShared;
      }

      @Override
      public String toString()
      {
        return path + " -- " + date + " shared:" + isShared;
      }
    }

    /**
     * Liefert eine {@link List} mit den aktuell auf dem System vorhandenen
     * WollMux-Installationen.
     * 
     * @author Christoph Lutz, Matthias Benkmann (D-III-ITD-D101)
     */
    private List<WollMuxInstallationDescriptor> getInstallations()
    {
      List<WollMuxInstallationDescriptor> wmInstallations =
        new ArrayList<WollMuxInstallationDescriptor>();

      // Installationspfade der Pakete bestimmen:
      String myPath = null; // user-Pfad
      String oooPath = null; // shared-Pfad
      String oooPathNew = null; // shared-Pfad (OOo 3.x)

      try
      {
        XStringSubstitution xSS =
          UNO.XStringSubstitution(UNO.createUNOService("com.sun.star.util.PathSubstitution"));

        // Benutzerinstallationspfad LiMux =
        // /home/<Benutzer>/.openoffice.org2/user
        // Benutzerinstallationspfad Windows 2000 C:/Dokumente und
        // Einstellungen/<Benutzer>/Anwendungsdaten/OpenOffice.org2/user
        myPath =
          xSS.substituteVariables("$(user)/uno_packages/cache/uno_packages/", true);
        // Sharedinstallationspfad LiMux /opt/openoffice.org2.0/
        // Sharedinstallationspfad Windows C:/Programme/OpenOffice.org<version>
        oooPath =
          xSS.substituteVariables("$(inst)/share/uno_packages/cache/uno_packages/",
            true);
        try
        {
          oooPathNew =
            xSS.substituteVariables(
              "$(brandbaseurl)/share/uno_packages/cache/uno_packages/", true);
        }
        catch (NoSuchElementException e)
        {
          // OOo 2.x does not have $(brandbaseurl)
        }

      }
      catch (java.lang.Exception e)
      {
        Logger.error(e);
        return wmInstallations;
      }

      if (myPath == null || oooPath == null)
      {
        Logger.error(L.m("Bestimmung der Installationspfade für das WollMux-Paket fehlgeschlagen."));
        return wmInstallations;
      }

      findWollMuxInstallations(wmInstallations, myPath, false);
      findWollMuxInstallations(wmInstallations, oooPath, true);
      if (oooPathNew != null)
        findWollMuxInstallations(wmInstallations, oooPathNew, true);

      return wmInstallations;
    }

    /**
     * Sucht im übergebenen Pfad path nach Verzeichnissen die WollMux.oxt enthalten
     * und fügt die Information zu wmInstallations hinzu.
     * 
     * @author Bettina Bauer, Christoph Lutz, Matthias Benkmann (D-III-ITD-D101)
     */
    private static void findWollMuxInstallations(
        List<WollMuxInstallationDescriptor> wmInstallations, String path,
        boolean isShared)
    {
      URI uriPath;
      uriPath = null;
      try
      {
        uriPath = new URI(path);
      }
      catch (URISyntaxException e)
      {
        Logger.error(e);
        return;
      }

      File[] installedPackages = new File(uriPath).listFiles();
      if (installedPackages != null)
      {
        // iterieren über die Installationsverzeichnisse mit automatisch
        // generierten Namen (z.B. 31GFBd_)
        for (int i = 0; i < installedPackages.length; i++)
        {
          if (installedPackages[i].isDirectory())
          {
            File dir = installedPackages[i];
            File[] dateien = dir.listFiles();
            for (int j = 0; j < dateien.length; j++)
            {
              // Wenn das Verzeichnis WollMux.oxt enthält, speichern des
              // Verzeichnisnames und des Verzeichnisdatum in einer HashMap
              if (dateien[j].isDirectory()
                && dateien[j].getName().startsWith("WollMux."))
              {
                // Name des Verzeichnis in dem sich WollMux.oxt befindet
                String directoryName = dateien[j].getAbsolutePath();
                // Datum des Verzeichnis in dem sich WollMux.oxt befindet
                Date directoryDate = new Date(dateien[j].lastModified());
                wmInstallations.add(new WollMuxInstallationDescriptor(directoryName,
                  directoryDate, isShared));
              }
            }
          }
        }
      }
    }

    @Override
    public String toString()
    {
      return this.getClass().getSimpleName() + "()";
    }
  }

  // *******************************************************************************************
  // Globale Helper-Methoden

  private static ConfigThingy requireLastSection(ConfigThingy cf, String sectionName)
      throws ConfigurationErrorException
  {
    try
    {
      return cf.query(sectionName).getLastChild();
    }
    catch (NodeNotFoundException e)
    {
      throw new ConfigurationErrorException(L.m(
        "Der Schlüssel '%1' fehlt in der Konfigurationsdatei.", sectionName), e);
    }
  }
}
