package de.muenchen.allg.itd51.wollmux.sidebar;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import com.sun.star.accessibility.XAccessible;
import com.sun.star.awt.WindowEvent;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogProvider;
import com.sun.star.awt.XSimpleTabController;
import com.sun.star.awt.XTabController;
import com.sun.star.awt.XTabControllerModel;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XUnoControlContainer;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.awt.tab.XTabPage;
import com.sun.star.awt.tab.XTabPageContainer;
import com.sun.star.awt.tab.XTabPageContainerModel;
import com.sun.star.awt.tab.XTabPageModel;
import com.sun.star.beans.NamedValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameContainer;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.ui.LayoutSize;
import com.sun.star.ui.XSidebarPanel;
import com.sun.star.ui.XToolPanel;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoProps;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlModel;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlModel.Align;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlModel.ControlType;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlModel.Orientation;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlProperties;
import de.muenchen.allg.itd51.wollmux.core.dialog.SimpleDialogLayout;
import de.muenchen.allg.itd51.wollmux.core.dialog.adapter.AbstractWindowListener;

public class FormGUISidebar implements XToolPanel, XSidebarPanel
{
  private XComponentContext context;
  private XWindow parentWindow;
  private XWindow window;
  private SimpleDialogLayout layout;
  private XWindowPeer windowPeer;
  private XToolkit toolkit;

  private AbstractWindowListener windowAdapter = new AbstractWindowListener()
  {
    @Override
    public void windowResized(WindowEvent e)
    {
      layout.draw();
    }
  };

  private XControlModel unoControlContainerModel = null;
  private XControl dialogControl = null;
  private XTabPageContainerModel xTabPageContainerModel;

  public FormGUISidebar(XComponentContext context, XWindow parentWindow)
  {
    this.context = context;
    this.parentWindow = parentWindow;
    this.parentWindow.addWindowListener(this.windowAdapter);
    // UNODialogFactory dialogFactory = new UNODialogFactory();

    Object cont = UNO.createUNOService("com.sun.star.awt.UnoControlContainer");
    dialogControl = UnoRuntime.queryInterface(XControl.class, cont);

    // Instanzierung eines ControlContainers für das aktuelle Fenster
    Object unoControlContainerModelO = UNO
        .createUNOService("com.sun.star.awt.UnoControlContainerModel");
    unoControlContainerModel = UnoRuntime.queryInterface(XControlModel.class,
        unoControlContainerModelO);
    dialogControl.setModel(unoControlContainerModel);

    XWindowPeer parentWindowPeer = UnoRuntime.queryInterface(XWindowPeer.class, this.parentWindow);

    if (parentWindowPeer != null)
    {
      dialogControl.createPeer(toolkit, parentWindowPeer);
      window = UNO.XWindow(dialogControl);
      layout = new SimpleDialogLayout(window);
      
      // TabControl muss zuerst in den Container eingefügt werden, bevor die TabPages hinzugefügt werden können
      layout.addControlsToList(addTabControl4());
      
      // Jetzt die TabPages einfügen
      XTabPageModel xTabPageModel1 = xTabPageContainerModel.createTabPage((short) 1);      
      xTabPageModel1.setTitle("Page1");
      xTabPageModel1.setEnabled(true);
      XTabPageModel xTabPageModel2 = xTabPageContainerModel.createTabPage((short) 2);
      xTabPageModel2.setTitle("Page2");
      xTabPageModel2.setEnabled(true);
      try {
        xTabPageContainerModel.insertByIndex(0, xTabPageModel1);
        xTabPageContainerModel.insertByIndex(1, xTabPageModel2);
      } catch (com.sun.star.lang.IllegalArgumentException | com.sun.star.lang.IndexOutOfBoundsException
            | WrappedTargetException e) {
        e.printStackTrace();
      }
      
      window.setVisible(true);
      window.setEnable(true);
    }

    layout.draw();
  }
  
  private void addCibTabTest()
  {

    // create the dialog
    Object oDialogProvider;
    try
    {
      oDialogProvider = UNO.xMCF.createInstanceWithContext("com.sun.star.awt.DialogProvider",
          UNO.defaultContext);

      XDialogProvider xDialogProvider = UnoRuntime.queryInterface(XDialogProvider.class,
          oDialogProvider);
      XDialog xDialog = xDialogProvider
          .createDialog("vnd.sun.star.script:Standard.Dialog1?location=application");

      XControlModel xDialogModel = UnoRuntime.queryInterface(XControl.class, xDialog).getModel();
      XMultiServiceFactory xMsf = UnoRuntime.queryInterface(XMultiServiceFactory.class,
          xDialogModel);
      XNameContainer xNameContainer = UnoRuntime.queryInterface(XNameContainer.class, xDialogModel);

      // create the tab pages container model
      Object tabPagesModel = UNO.xMSF
          .createInstance("com.sun.star.awt.tab.UnoControlTabPageContainerModel");
      xNameContainer.insertByName("tab", tabPagesModel);
      XPropertySet xPropSet = UnoRuntime.queryInterface(XPropertySet.class, tabPagesModel);

      //xPropSet.setPropertyValue("Width", 100);
      //xPropSet.setPropertyValue("Height", 100);

      XTabPageContainerModel xTabPagesModel = UnoRuntime
          .queryInterface(XTabPageContainerModel.class, tabPagesModel);

      // add the first page
//      XTabPageModel xTabPageModel1 = xTabPagesModel.createTabPage((short) 1);
//      xTabPageModel1.setTitle("Page 1");
//      xTabPagesModel.insertByIndex(0, xTabPageModel1);
//
//      // add the second page
//      XTabPageModel xTabPageModel2 = xTabPagesModel.createTabPage((short) 2);
//      xTabPageModel2.setTitle("Page 2");
//      xTabPagesModel.insertByIndex(1, xTabPageModel2);
      
      for (int i = 1; i < 15; i++) {
        XTabPageModel xTabPageModel2 = xTabPagesModel.createTabPage((short) i);
        xTabPageModel2.setTitle("Page " + i);
        xTabPagesModel.insertByIndex(i - 1, xTabPageModel2);
      }
      
      // execute the dialog
      xDialog.execute();
    } catch (Exception e)
    {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }

  }

  private ControlModel addTabControl4()
    {
      List<ControlProperties> tabControls = new ArrayList<>();
      Object tabPageContainerModel = UNO.createUNOService("com.sun.star.awt.tab.UnoControlTabPageContainerModel");
      xTabPageContainerModel = UnoRuntime
          .queryInterface(XTabPageContainerModel.class, tabPageContainerModel);
  
      Object tabControl = UNO.createUNOService("com.sun.star.awt.tab.UnoControlTabPageContainer");
      XControl xControl = (XControl)UnoRuntime.queryInterface(XControl.class, tabControl );
      
      XControlModel xControlModel = (XControlModel)UnoRuntime.queryInterface(
              XControlModel.class, tabPageContainerModel); 
      xControl.setModel(xControlModel);
      Object toolkit = UNO.createUNOService("com.sun.star.awt.Toolkit"); 
      XToolkit xToolkit = (XToolkit)UnoRuntime.queryInterface(XToolkit.class, toolkit);
      xControl.createPeer(xToolkit, dialogControl.getPeer());
      
  
      ControlProperties controlProperties = new ControlProperties(ControlType.EDIT, "tabContainer");
      controlProperties.setControlPercentSize(100, 50);
      controlProperties.setXControl(xControl);
      
      tabControls.add(controlProperties);
  
      return new ControlModel(Orientation.HORIZONTAL, Align.NONE, tabControls, Optional.empty());
    }

  private ControlModel addTabControl2()
  {
    List<ControlProperties> tabControls = new ArrayList<>();

    try
    {
      XControlModel containerModel = dialogControl.getModel();
      XControlContainer container = layout.getControlContainer();

      Object multiPageModel = UNO.createUNOService("com.sun.star.awt.UnoMultiPageModel");
      XControl xMultiPageModelControl = UnoRuntime.queryInterface(XControl.class, multiPageModel); // null
      container.addControl("tab", xMultiPageModelControl);
      XPropertySet xProp = UnoRuntime.queryInterface(XPropertySet.class, multiPageModel);
      xProp.setPropertyValue("Width", 100);
      xProp.setPropertyValue("Height", 100);

      XSimpleTabController xTabController = UnoRuntime.queryInterface(XSimpleTabController.class,
          xMultiPageModelControl);

      Object pageModel1 = UNO.createUNOService("com.sun.star.awt.UnoPageModel");
      XControl xPageModel1 = UnoRuntime.queryInterface(XControl.class,
          "com.sun.star.awt.UnoPageModel");
      container.addControl("page1", xPageModel1);
      NamedValue nv = new NamedValue("Title", "Page1");
      xTabController.setTabProps(1, new NamedValue[] { nv });

      Object pageModel2 = UNO.createUNOService("com.sun.star.awt.UnoPageModel");
      XControl xPageModel2 = UnoRuntime.queryInterface(XControl.class,
          "com.sun.star.awt.UnoPageModel");
      container.addControl("page2", xPageModel2);
      NamedValue nv2 = new NamedValue("Title", "Page2");
      xTabController.setTabProps(2, new NamedValue[] { nv2 });

      XControl xControlTabController = UnoRuntime.queryInterface(XControl.class, xTabController);
      ControlProperties controlProperties = new ControlProperties(ControlType.EDIT, "tabContainer");
      controlProperties.setXControl(xMultiPageModelControl);

      tabControls.add(controlProperties);
    } catch (Exception ex)
    {
      //
    }

    return new ControlModel(Orientation.HORIZONTAL, Align.NONE, tabControls, Optional.empty());
  }
  
  private ControlModel addTabControl5()
  {

    List<ControlProperties> tabControls = new ArrayList<>();

    try
    {
      Object multiPageModel = UNO.createUNOService("com.sun.star.awt.UnoPageControl"); //UnoMultiPageControl oder UnoMultiPageModel?
      XNameContainer nameContainer = UnoRuntime.queryInterface(XNameContainer.class, multiPageModel);
      
      //XControl xMultiPageModelControl = UnoRuntime.queryInterface(XControl.class, multiPageModel); // null
//      XPropertySet xProp = UnoRuntime.queryInterface(XPropertySet.class, multiPageModel);
//      xProp.setPropertyValue("Width", 100);
//      xProp.setPropertyValue("Height", 100);

      XSimpleTabController xTabController = UnoRuntime.queryInterface(XSimpleTabController.class,
          multiPageModel); // eigentlich sollte hier multiPageModel -> XControl cast übergeben werden
      int newTabId = xTabController.insertTab();
      int newTabId1 = xTabController.insertTab();
      
      NamedValue nv = new NamedValue("Title", "Page1");
      xTabController.setTabProps(newTabId, new NamedValue[] { nv });
      
      NamedValue nv1 = new NamedValue("Title", "Page2");
      xTabController.setTabProps(newTabId1, new NamedValue[] { nv1 });
      
      nameContainer.insertByName("tabcontrol", multiPageModel);

      XControl xControlTabController = UnoRuntime.queryInterface(XControl.class, nameContainer);
      ControlProperties controlProperties = new ControlProperties(ControlType.EDIT, "tabContainer");
      controlProperties.setXControl(xControlTabController);

      tabControls.add(controlProperties);
    } catch (Exception ex)
    {
      //
    }

    return new ControlModel(Orientation.HORIZONTAL, Align.NONE, tabControls, Optional.empty());
  }

  private void addTabControl()
  {
    List<ControlProperties> tabControls = new ArrayList<>();

    ControlProperties BTN = new ControlProperties(ControlType.EDIT, "tabContainer");
    BTN.setControlPercentSize(100, 100);

    Object oTabPageContainerModel = UNO
        .createUNOService("com.sun.star.awt.tab.UnoControlTabPageContainerModel");
    XTabPageContainerModel tabPageContainerModel = UnoRuntime
        .queryInterface(XTabPageContainerModel.class, oTabPageContainerModel);

    Object oTabPageContainer = UNO
        .createUNOService("com.sun.star.awt.tab.UnoControlTabPageContainer");
    XTabPageContainer tabPageContainer = UnoRuntime.queryInterface(XTabPageContainer.class,
        oTabPageContainer);

    XControlModel controlModel = UnoRuntime.queryInterface(XControlModel.class,
        tabPageContainerModel);
    UNO.XControl(tabPageContainer).setModel(controlModel);
    XTabPageModel tab1 = tabPageContainerModel.createTabPage((short) 1);
    XTabPageModel tab2 = tabPageContainerModel.createTabPage((short) 2);

    tab1.setTitle("Test1");
    tab1.setEnabled(true);

    tab2.setTitle("Test2");
    tab2.setEnabled(true);

    UnoProps newDesc = new UnoProps();
    newDesc.setPropertyValue("TabPage", tab1);

    
    XPropertySet propSet = UNO.XPropertySet(tab1);
     try
     {
     // tabPageContainerModel.loadTabPage((short) 11,
     // "vnd.sun.star.script:WollMux.email_auth?location=application");
    
       XControlContainer xContainer = UnoRuntime.queryInterface(XControlContainer.class, tabPageContainer);
     tabPageContainerModel.insertByIndex(0, tab1);
     tabPageContainerModel.insertByIndex(1, tab2);
     } catch (IllegalArgumentException e)
     {
     // TODO Auto-generated catch block
     e.printStackTrace();
     } catch (IndexOutOfBoundsException e)
     {
     // TODO Auto-generated catch block
     e.printStackTrace();
     } catch (com.sun.star.lang.IllegalArgumentException e)
     {
     // TODO Auto-generated catch block
     e.printStackTrace();
     } catch (com.sun.star.lang.IndexOutOfBoundsException e)
     {
     // TODO Auto-generated catch block
     e.printStackTrace();
     } catch (WrappedTargetException e)
    {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
     
    Object tabController = UNO.createUNOService("com.sun.star.awt.TabController");
    XTabController xTabController = UnoRuntime.queryInterface(XTabController.class, tabController);

    Object tabControllerModel = UNO.createUNOService("com.sun.star.awt.TabControllerModel");
    XTabControllerModel tabCModel = UnoRuntime.queryInterface(XTabControllerModel.class,
        tabControllerModel);
    xTabController.setModel(tabCModel);
    xTabController.setContainer(UnoRuntime.queryInterface(XControlContainer.class, tabPageContainer));
    // tabCModel.setControlModels(new XControlModel[] { xControl.getModel() });

    // XControl xUnoCC = null;
    XControl xunoCContainer = null;

      // Object unoControlContainer =
      // UNO.xMSF.createInstance("com.sun.star.awt.UnoControlContainer");
      // xUnoCC = UnoRuntime.queryInterface(XControl.class, unoControlContainer);
      // xUnoCC.setModel(xControl.getModel());

      // XUnoControlContainer xcc = UnoRuntime.queryInterface(XUnoControlContainer.class,
      // unoControlContainer);
    //

    // could work
    XUnoControlContainer xcc = UnoRuntime.queryInterface(XUnoControlContainer.class,
        layout.getControlContainer());
    xcc.addTabController(xTabController);
    xcc.getTabControllers(); //

    BTN.setXControl(UNO.XControl(tabPageContainerModel));
    tabControls.add(BTN);
    layout.addControlsToList(
        new ControlModel(Orientation.HORIZONTAL, Align.NONE, tabControls, Optional.empty()));

      // XControlContainer cont = UnoRuntime.queryInterface(XControlContainer.class, xcc);
      // layout.setControlContainer(cont);
    // xTabController.setContainer(UnoRuntime.queryInterface(XControlContainer.class, xcc));
    // xunoCContainer = UnoRuntime.queryInterface(XControl.class, xcc);


    // ControlProperties controlProperties = new ControlProperties(ControlType.EDIT,
    // "tabContainer");
    // controlProperties.setXControl(xunoCContainer);

    // tabControls.add(abortBtn);

    // tabControls.add(controlProperties);

    // return new ControlModel(Orientation.HORIZONTAL, Align.NONE, tabControls, Optional.empty());
  }

  private ControlModel addBottomControls()
  {
    List<ControlProperties> bottomControls = new ArrayList<>();

    ControlProperties abortBtn = new ControlProperties(ControlType.BUTTON, "abortBtn");
    abortBtn.setControlPercentSize(50, 30);
    abortBtn.setLabel("Abbrechen");
    
    // UnoRuntime.queryInterface(XButton.class, abortBtn.getValue())
    // .addActionListener(abortBtnActionListener);

    bottomControls.add(abortBtn);

    return new ControlModel(Orientation.HORIZONTAL, Align.NONE, bottomControls, Optional.empty());
  }

  @Override
  public LayoutSize getHeightForWidth(int arg0)
  {
    // int height = layout.getHeight();

    return new LayoutSize(300, 100, 100);
  }

  @Override
  public int getMinimalWidth()
  {
    return 100;
  }

  @Override
  public XAccessible createAccessible(XAccessible arg0)
  {
    if (window == null)
    {
      //
    }

    return UnoRuntime.queryInterface(XAccessible.class, getWindow());
  }

  @Override
  public XWindow getWindow()
  {
    if (window == null)
    {
      throw new DisposedException("", this);
    }
    return window;
  }

}
