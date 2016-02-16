package de.muenchen.allg.itd51.wollmux.dialog.formmodel;

import java.awt.event.ActionListener;
import java.util.HashMap;
import java.util.Map;
import java.util.Vector;

import com.sun.star.frame.XFrame;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.itd51.parser.ConfigThingy;
import de.muenchen.allg.itd51.parser.NodeNotFoundException;
import de.muenchen.allg.itd51.wollmux.L;
import de.muenchen.allg.itd51.wollmux.Logger;
import de.muenchen.allg.itd51.wollmux.TextDocumentModel;
import de.muenchen.allg.itd51.wollmux.WollMuxFiles;
import de.muenchen.allg.itd51.wollmux.db.MailMergeDatasource;
import de.muenchen.allg.itd51.wollmux.dialog.DialogLibrary;
import de.muenchen.allg.itd51.wollmux.dialog.FormGUI;
import de.muenchen.allg.itd51.wollmux.event.Dispatch;
import de.muenchen.allg.itd51.wollmux.event.WollMuxEventHandler;
import de.muenchen.allg.itd51.wollmux.func.FunctionLibrary;

/**
 * Repräsentiert ein FormModel für ein einfaches Formular mit genau einem
 * zugehörigen Formulardokument. Diese Klasse sorgt als Wrapper im Wesentlichen nur
 * dafür, dass alle Methodenaufrufe des FormModels in die ensprechenden
 * WollMuxEvents verpackt werden und somit korrekt synchronisiert ausgeführt
 * werden.
 */
public class SingleDocumentFormModel implements FormModel
{
  private final TextDocumentModel doc;

  private final ConfigThingy formFensterConf;

  private final ConfigThingy formConf;

  private final Map<Object, Object> functionContext;

  private final FunctionLibrary funcLib;

  private final DialogLibrary dialogLib;

  private final boolean visible;

  private final String defaultWindowAttributes;

  private FormGUI formGUI = null;

  /**
  * vFormGUIs beinhaltet die Referenzen der z.Z. ausgeführten FormGUIs
  */
  public static Vector<FormGUI> vFormGUIs=new Vector<FormGUI>();

  /**
  * vFrames beinhaltet die Referenzen der z.Z. ausgeführten Frames mit FormGUI
  */
  public static Vector<XFrame> vFrames=new Vector<XFrame>();

  /**
   * Konstruktor für ein SingleDocumentFormModel mit dem zugehörigen
   * TextDocumentModel doc.
   * 
   * @param doc
   *          Das zugeordnete TextDocumentModel.
   * @param formFensterConf
   *          Der Formular-Unterabschnitt des Fenster-Abschnitts von wollmux.conf
   *          (wird für createFormGUI() benötigt).
   * @param formConf
   *          der Formular-Knoten, der die Formularbeschreibung enthält (wird für
   *          createFormGUI() benötigt).
   * @param functionContext
   *          der Kontext für Funktionen, die einen benötigen (wird für
   *          createFormGUI() benötigt).
   * @param funcLib
   *          die Funktionsbibliothek, die zur Auswertung von Trafos, Plausis etc.
   *          herangezogen werden soll.
   * @param dialogLib
   *          die Dialogbibliothek, die die Dialoge bereitstellt, die für
   *          automatisch zu befüllende Formularfelder benötigt werden (wird für
   *          createFormGUI() benötigt).
   * @param visible
   *          false zeigt an, dass die FormGUI unsichtbar sein soll.
   */
  public SingleDocumentFormModel(final TextDocumentModel doc,
      final ConfigThingy formFensterConf, final ConfigThingy formConf,
      final Map<Object, Object> functionContext, final FunctionLibrary funcLib,
      final DialogLibrary dialogLib, boolean visible)
  {
    this.doc = doc;
    this.formFensterConf = formFensterConf;
    this.formConf = formConf;
    this.functionContext = functionContext;
    this.funcLib = funcLib;
    this.dialogLib = dialogLib;
    this.visible = visible;

    // Standard-Fensterattribute vor dem Start der Form-GUI sichern um nach
    // dem Schließen des Formulardokuments die Standard-Werte wieder
    // herstellen zu können. Die Standard-Attribute ändern sich (OOo-seitig)
    // immer dann, wenn ein Dokument (mitsamt Fenster) geschlossen wird. Dann
    // merkt sich OOo die Position und Größe des zuletzt geschlossenen
    // Fensters.
    if (visible)
      this.defaultWindowAttributes = getDefaultWindowAttributes();
    else
      this.defaultWindowAttributes = null;
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#close()
   */
  @Override
  public void close()
  {
    WollMuxEventHandler.handleCloseTextDocument(doc);
  }

  @Override
  public void closeAndOpenExt(String ext)
  {
    WollMuxEventHandler.handleCloseAndOpenExt(doc, ext);
  }

  @Override
  public void saveTempAndOpenExt(String ext)
  {
    WollMuxEventHandler.handleSaveTempAndOpenExt(doc, ext);
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#setWindowVisible(boolean)
   */
  @Override
  public void setWindowVisible(boolean vis)
  {
    /*
     * Einmal unsichtbar, immer unsichtbar. Weiß nicht, ob das sinnvoll ist, aber
     * die ganze Methode wird soweit ich sehen kann derzeit nicht verwendet, also
     * ist es egal. Falls es hier erlaubt wird, das Fenster sichtbar zu schalten,
     * dann müsste noch einiges anderes geändert werden, z.B. müsste die FormGUI
     * sichtbar werden.
     */
    if (visible) WollMuxEventHandler.handleSetWindowVisible(doc, vis);
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#setWindowPosSize(int, int, int,
   * int)
   */
  @Override
  public void setWindowPosSize(int docX, int docY, int docWidth, int docHeight)
  {
    if (visible)
      WollMuxEventHandler.handleSetWindowPosSize(doc, docX, docY, docWidth,
        docHeight);
  }

  /*
   * (non-Javadoc)
   * 
   * @see
   * de.muenchen.allg.itd51.wollmux.FormModel#setVisibleState(java.lang.String,
   * boolean)
   */
  @Override
  public void setVisibleState(String groupId, boolean visible)
  {
    WollMuxEventHandler.handleSetVisibleState(doc, groupId, visible, null);
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#valueChanged(java.lang.String,
   * java.lang.String)
   */
  @Override
  public void valueChanged(String fieldId, String newValue)
  {
    if (fieldId.length() > 0)
      WollMuxEventHandler.handleFormValueChanged(doc, fieldId, newValue);
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#focusGained(java.lang.String)
   */
  @Override
  public void focusGained(String fieldId)
  {
    if (visible) WollMuxEventHandler.handleFocusFormField(doc, fieldId);
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#focusLost(java.lang.String)
   */
  @Override
  public void focusLost(String fieldId)
  {}

  /*
   * (non-Javadoc)
   * 
   * @see
   * de.muenchen.allg.itd51.wollmux.FormModel#disposed(de.muenchen.allg.itd51.wollmux
   * .TextDocumentModel)
   */
  @Override
  public void disposing(TextDocumentModel source)
  {
    if (doc.equals(source))
    {
      if (formGUI != null)
      {
        formGUI.dispose();
        formGUI = null;
      }

      // Rücksetzen des defaultWindowAttributes auf den Wert vor dem Schließen
      // des Formulardokuments.
      if (defaultWindowAttributes != null)
        setDefaultWindowAttributes(defaultWindowAttributes);
    }
  }

  /**
   * Diese Hilfsmethode liest das Attribut ooSetupFactoryWindowAttributes aus dem
   * Konfigurationsknoten
   * "/org.openoffice.Setup/Office/Factories/com.sun.star.text.TextDocument" der
   * OOo-Konfiguration, welches die Standard-FensterAttribute enthält, mit denen
   * neue Fenster für TextDokumente erzeugt werden.
   * 
   * @return
   */
  private static String getDefaultWindowAttributes()
  {
    try
    {
      Object cp =
        UNO.createUNOService("com.sun.star.configuration.ConfigurationProvider");

      // creation arguments: nodepath
      com.sun.star.beans.PropertyValue aPathArgument =
        new com.sun.star.beans.PropertyValue();
      aPathArgument.Name = "nodepath";
      aPathArgument.Value =
        "/org.openoffice.Setup/Office/Factories/com.sun.star.text.TextDocument";
      Object[] aArguments = new Object[1];
      aArguments[0] = aPathArgument;

      Object ca =
        UNO.XMultiServiceFactory(cp).createInstanceWithArguments(
          "com.sun.star.configuration.ConfigurationAccess", aArguments);

      return UNO.getProperty(ca, "ooSetupFactoryWindowAttributes").toString();
    }
    catch (java.lang.Exception e)
    {}
    return null;
  }

  /**
   * Diese Hilfsmethode setzt das Attribut ooSetupFactoryWindowAttributes aus dem
   * Konfigurationsknoten
   * "/org.openoffice.Setup/Office/Factories/com.sun.star.text.TextDocument" der
   * OOo-Konfiguration auf den neuen Wert value, der (am besten) über einen
   * vorhergehenden Aufruf von getDefaultWindowAttributes() gewonnen wird.
   * 
   * @param value
   */
  private static void setDefaultWindowAttributes(String value)
  {
    try
    {
      Object cp =
        UNO.createUNOService("com.sun.star.configuration.ConfigurationProvider");

      // creation arguments: nodepath
      com.sun.star.beans.PropertyValue aPathArgument =
        new com.sun.star.beans.PropertyValue();
      aPathArgument.Name = "nodepath";
      aPathArgument.Value =
        "/org.openoffice.Setup/Office/Factories/com.sun.star.text.TextDocument";
      Object[] aArguments = new Object[1];
      aArguments[0] = aPathArgument;

      Object ca =
        UNO.XMultiServiceFactory(cp).createInstanceWithArguments(
          "com.sun.star.configuration.ConfigurationUpdateAccess", aArguments);

      UNO.setProperty(ca, "ooSetupFactoryWindowAttributes", value);

      UNO.XChangesBatch(ca).commitChanges();
    }
    catch (java.lang.Exception e)
    {
      Logger.error(e);
    }
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#controllerInitCompleted()
   */
  @Override
  public void formControllerInitCompleted()
  {
    WollMuxEventHandler.handleFormControllerInitCompleted(doc);
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#print()
   */
  @Override
  public void print()
  {
    UNO.dispatch(doc.doc, Dispatch.DISP_unoPrint);
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#pdf()
   */
  @Override
  public void pdf()
  {
    UNO.dispatch(doc.doc, ".uno:ExportToPDF");
  }

  @Override
  public void save()
  {
    UNO.dispatch(doc.doc, ".uno:Save");
  }

  @Override
  public void saveAs()
  {
    UNO.dispatch(doc.doc, ".uno:SaveAs");
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#setValue(java.lang.String,
   * java.lang.String, java.awt.event.ActionListener)
   */
  @Override
  public void setValue(String fieldId, String value, ActionListener listener)
  {
    if (formGUI != null)
      formGUI.getController().setValue(fieldId, value, listener);
  }

  /*
   * (non-Javadoc)
   * 
   * @see de.muenchen.allg.itd51.wollmux.FormModel#startFormGUI()
   */
  @Override
  public void startFormGUI()
  {
    boolean containsFrames=false;
    containsFrames=SingleDocumentFormModel.vFrames.contains(doc.getFrame());

    //Schaue ob bereits eine Instanz genau dieses Formulars geöffnet ist, falls ja wird das nun ungültige FormGUI beendet 
    if (containsFrames){  
      //Hole Index des ungültigen FormGUI
      int frameIndex=SingleDocumentFormModel.vFrames.indexOf(doc.getFrame());
      
      SingleDocumentFormModel.vFormGUIs.get(frameIndex).dispose();
      SingleDocumentFormModel.vFormGUIs.remove(frameIndex);
      SingleDocumentFormModel.vFrames.remove(frameIndex);
      Logger.debug(L.m("FormGUI an der Stelle %1 beendet.", frameIndex));
    }
    
    HashMap<String, String> idToPresetValue = doc.getIDToPresetValue();
    formGUI =
      new FormGUI(formFensterConf, formConf, this, idToPresetValue,
        functionContext, funcLib, dialogLib, visible);
    
    // füge FormGUI Refenrenz und die dazugehörigen Frames zu den Klassenvariable hinzu
    SingleDocumentFormModel.vFormGUIs.add(formGUI);
    SingleDocumentFormModel.vFrames.add(doc.getFrame());  
  }

  @Override
  public String getWindowTitle()
  {
    try
    {
      XFrame frame = UNO.XModel(doc.doc).getCurrentController().getFrame();
      String frameTitle = (String) UNO.getProperty(frame, "Title");
      frameTitle = MailMergeDatasource.stripOpenOfficeFromWindowName(frameTitle);
      return frameTitle;
    }
    catch (Exception x)
    {
      return null;
    }
  }

  /**
   * Erzeugt ein FormModel für ein einfaches Formular mit genau einem zugehörigen
   * Formulardokument.
   * 
   * @param doc
   *          Das Dokument zu dem ein FormModel erzeugt werden soll.
   * @param visible
   *          false zeigt an, dass das Dokument unsichtbar ist (und es die FormGUI
   *          auch sein sollte).
   * @return ein FormModel dem genau ein Formulardokument zugeordnet ist.
   * @throws InvalidFormDescriptorException
   */
  public static FormModel createSingleDocumentFormModel(TextDocumentModel doc,
      boolean visible) throws InvalidFormDescriptorException
  {
  
    // Abschnitt "Formular" holen:
    ConfigThingy formConf;
    try
    {
      formConf = doc.getFormDescription().get("Formular");
    }
    catch (NodeNotFoundException e)
    {
      throw new InvalidFormDescriptorException(
        L.m("Kein Abschnitt 'Formular' in der Formularbeschreibung vorhanden"));
    }
  
    // Abschnitt Fenster/Formular aus wollmuxConf holen:
    ConfigThingy formFensterConf;
    try
    {
      formFensterConf =
        WollMuxFiles.getWollmuxConf().query("Fenster").query("Formular").getLastChild();
    }
    catch (NodeNotFoundException x)
    {
      formFensterConf = new ConfigThingy("");
    }
  
    return new SingleDocumentFormModel(doc, formFensterConf, formConf,
      doc.getFunctionContext(), doc.getFunctionLibrary(), doc.getDialogLibrary(),
      visible);
  }
}