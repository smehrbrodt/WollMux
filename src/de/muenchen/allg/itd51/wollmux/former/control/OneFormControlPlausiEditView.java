/*
* Dateiname: OneFormControlPlausiEditView.java
* Projekt  : WollMux
* Funktion : Stellt das PLAUSI-Attribut eines FormControlModels dar und erlaubt seine Bearbeitung.
* 
* Copyright: Landeshauptstadt M�nchen
*
* �nderungshistorie:
* Datum      | Wer | �nderungsgrund
* -------------------------------------------------------------------
* 24.10.2006 | BNK | Erstellung
* -------------------------------------------------------------------
*
* @author Matthias Benkmann (D-III-ITD 5.1)
* @version 1.0
* 
*/
package de.muenchen.allg.itd51.wollmux.former.control;

import de.muenchen.allg.itd51.wollmux.former.function.FunctionSelectionAccessView;
import de.muenchen.allg.itd51.wollmux.former.view.ViewChangeListener;
import de.muenchen.allg.itd51.wollmux.func.FunctionLibrary;

/**
 * Stellt das PLAUSI-Attribut eines FormControlModels dar und erlaubt seine Bearbeitung.
 *
 * @author Matthias Benkmann (D-III-ITD 5.1)
 */
public class OneFormControlPlausiEditView extends FunctionSelectionAccessView
{
  /**
   * Typischerweise ein Container, der die View enth�lt und daher �ber �nderungen
   * auf dem Laufenden gehalten werden muss.
   */
  private ViewChangeListener bigDaddy;
  
  /**
   * Das Model zu dieser View.
   */
  private FormControlModel model;
  
  /**
   * Erzeugt eine neue View.
   * @param model das Model dessen Daten angezeigt werden sollen.
   * @param funcLib die Funktionsbibliothek deren Funktionen zur Verf�gung gestellt werden sollen.
   * @param myViewChangeListener typischerweise ein Container, der diese View enth�lt und �ber
   *        �nderungen informiert werden soll.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   * TESTED
   */
  public OneFormControlPlausiEditView(FormControlModel model, FunctionLibrary funcLib, ViewChangeListener bigDaddy)
  {
    super(model.getPlausiAccess(), funcLib);
    this.model = model;
    this.bigDaddy = bigDaddy;
    model.addListener(new MyModelChangeListener());
  }
  
  private class MyModelChangeListener implements FormControlModel.ModelChangeListener
  {
    public void modelRemoved(FormControlModel model)
    {
      if (bigDaddy != null)
        bigDaddy.viewShouldBeRemoved(OneFormControlPlausiEditView.this);
    }

    public void attributeChanged(FormControlModel model, int attributeId, Object newValue)
    {
    }
  }
  
  /**
   * Liefert das Model zu dieser View.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public FormControlModel getModel()
  {
    return model;
  }

}