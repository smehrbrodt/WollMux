package de.muenchen.allg.itd51.wollmux.sidebar;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.AbstractMap.SimpleEntry;

import com.sun.star.accessibility.XAccessible;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.WindowEvent;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
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
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.ui.LayoutSize;
import com.sun.star.ui.XSidebarPanel;
import com.sun.star.ui.XToolPanel;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.itd51.wollmux.core.constants.XButtonProperties;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlModel;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlProperties;
import de.muenchen.allg.itd51.wollmux.core.dialog.SimpleDialogLayout;
import de.muenchen.allg.itd51.wollmux.core.dialog.UNODialogFactory;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlModel.Align;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlModel.ControlType;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlModel.Dock;
import de.muenchen.allg.itd51.wollmux.core.dialog.ControlModel.Orientation;
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
      window.setEnable(true);
      window.setVisible(true);

      this.parentWindow.setVisible(true);

      layout = new SimpleDialogLayout(window);

      // layout.addControlsToList(addBottomControls());
      layout.addControlsToList(addTabControl5());

      window.setVisible(true);
    }

    layout.draw();

  }

  private ControlModel addTabControl4()
  {
    List<ControlProperties> tabControls = new ArrayList<>();

    Object tabPagesModel = UNO.createUNOService("com.sun.star.awt.tab.UnoControlTabPage");
    XTabPage xTabPage = UnoRuntime.queryInterface(XTabPage.class, tabPagesModel);
    XTabPageContainer xTabPageContainer = UnoRuntime.queryInterface(XTabPageContainer.class, xTabPage);
    XControl xControl = UnoRuntime.queryInterface(XControl.class, tabPagesModel);

    ControlProperties controlProperties = new ControlProperties(ControlType.EDIT, "tabContainer");
    controlProperties.setXControl(xControl);

    tabControls.add(controlProperties);

    return new ControlModel(Orientation.HORIZONTAL, Align.NONE, tabControls, Optional.empty());
  }

  private ControlModel addTabControl3()
  {
    List<ControlProperties> tabControls = new ArrayList<>();

    try
    {
      XControlModel containerModel = dialogControl.getModel();
      XControlContainer container = layout.getControlContainer();

      Object tabPagesModel = UNO
          .createUNOService("com.sun.star.awt.tab.UnoControlTabPageContainerModel");
      XTabPageContainerModel xTabPageContainerModel = UnoRuntime
          .queryInterface(XTabPageContainerModel.class, tabPagesModel);
      // XTabPageContainer ccp = UnoRuntime.queryInterface(XTabPageContainer.class,
      // xTabPageContainerModel);

      // page1
      XTabPageModel xTabPageModel1 = xTabPageContainerModel.createTabPage((short) 1);
      xTabPageModel1.setTitle("Page1");
      xTabPageModel1.setEnabled(true);
      xTabPageContainerModel.insertByIndex(0, xTabPageModel1);

      // page2
      XTabPageModel xTabPageModel2 = xTabPageContainerModel.createTabPage((short) 2);
      xTabPageModel2.setTitle("Page2");
      xTabPageModel2.setEnabled(true);
      xTabPageContainerModel.insertByIndex(1, xTabPageModel2);

      XControlModel xControlModel = UnoRuntime.queryInterface(XControlModel.class, tabPagesModel);

      Object tabPageContainer = UNO
          .createUNOService("com.sun.star.awt.tab.UnoControlTabPageContainer");
      XControl xTabPageControl = UnoRuntime.queryInterface(XControl.class, tabPageContainer);
      xTabPageControl.setModel(xControlModel); //UnknownPropertyException in Job.java

      ControlProperties controlProperties = new ControlProperties(ControlType.EDIT, "tabContainer");
      controlProperties.setXControl(xTabPageControl);

      tabControls.add(controlProperties);
    } catch (Exception ex)
    {

    }

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

  private ControlModel addTabControl()
  {
    List<ControlProperties> tabControls = new ArrayList<>();

    ControlProperties BTN = new ControlProperties(ControlType.EDIT, "tabContainer");
    BTN.setControlPercentSize(50, 30);

    // SimpleEntry<ControlProperties, XControl> abortBtn = layout
    // .convertToXControl(new ControlProperties(ControlType.TAB_CONTAINER, "tabContainer", 0, 30,
    // 50, 0,
    // new SimpleEntry<String[], Object[]>(new String[] { },
    // new Object[] { })));

    // Object control =
    // UNO.createUNOService("com.sun.star.awt.tab.UnoControlTabPageContainerModel");
    // XTabPageContainerModel tabPageContainer =
    // UnoRuntime.queryInterface(XTabPageContainerModel.class, control);
    //
    // XTabPageModel tabPageModel = tabPageContainer.createTabPage((short)0);
    // tabPageModel.setTitle("Test");
    // tabPageModel.setEnabled(true);
    //
    // XControl xControl = UnoRuntime.queryInterface(XControl.class, tabPageContainer);

    Object control = UNO.createUNOService("com.sun.star.awt.tab.UnoControlTabPage");
    XTabPage tabPage = UnoRuntime.queryInterface(XTabPage.class, control);
    XControl xControl = UnoRuntime.queryInterface(XControl.class, tabPage);

    Object tabController = UNO.createUNOService("com.sun.star.awt.TabController");
    XTabController xTabController = UnoRuntime.queryInterface(XTabController.class, tabController);

    Object tabControllerModel = UNO.createUNOService("com.sun.star.awt.TabControllerModel");
    XTabControllerModel tabCModel = UnoRuntime.queryInterface(XTabControllerModel.class,
        tabControllerModel);
    tabCModel.setControlModels(new XControlModel[] { BTN.getXControl().getModel() });

    // XControl xUnoCC = null;
    XControl xunoCContainer = null;
    try
    {
      Object unoControlContainer = UNO.xMSF.createInstance("com.sun.star.awt.UnoControlContainer");
      // xUnoCC = UnoRuntime.queryInterface(XControl.class, unoControlContainer);
      // xUnoCC.setModel(xControl.getModel());

      XUnoControlContainer xcc = UnoRuntime.queryInterface(XUnoControlContainer.class,
          unoControlContainer);
      xcc.addTabController(xTabController);

      // XControlContainer cont = UnoRuntime.queryInterface(XControlContainer.class, xcc);
      // layout.setControlContainer(cont);
      xTabController.setContainer(UnoRuntime.queryInterface(XControlContainer.class, xcc));
      xunoCContainer = UnoRuntime.queryInterface(XControl.class, xcc);

    } catch (Exception e)
    {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }

    ControlProperties controlProperties = new ControlProperties(ControlType.EDIT, "tabContainer");
    controlProperties.setXControl(xunoCContainer);

    // tabControls.add(abortBtn);

    tabControls.add(controlProperties);

    return new ControlModel(Orientation.HORIZONTAL, Align.NONE, tabControls, Optional.empty());
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
