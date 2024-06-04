using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ENAMI.Utility;
using DALPresupuestoAlerta;
using EntityPresupuestoAlerta;
using System.Data;
using AmbienteENAMI;
using DevExpress.Web;
using DevExpress.Web.ASPxTreeList;
using DevExpress.XtraPivotGrid;
using System.Drawing;
using DevExpress.Web.ASPxPivotGrid;
using System.Collections.Specialized;
using System.Data.SqlClient;
using DevExpress.Data;
using DevExpress.Spreadsheet;
using System.IO;

public partial class PresupuestoReportBD : System.Web.UI.Page
{
    const int permisoId = 3;//Ingreso Presupuesto.
  
    private int versionSession
    {
        get { return (int)Session["VersionSession"]; }
        set { Session["VersionSession"] = value; }
    }

    private int periodoSession
    {
        get { return (int)Session["PeriodoSession"]; }
        set { Session["PeriodoSession"] = value; }
    }

    private string cebeSession
    {
        get { return (string)Session["CEBESession"]; }
        set { Session["CEBESession"] = value; }
    }

    private string cecoSession
    {
        get { return (string)Session["CECOSession"]; }
        set { Session["CECOSession"] = value; }
    }

    private DataTable oDataTemp
    {
        get { return (DataTable)Session["DataTempPR"]; }
        set { Session["DataTempPR"] = value; }
    }

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!Page.IsPostBack)
        {
            Base oBase = new Base();
            if (new FuncionarioPermisoDAL(oBase.StrConexionPresupuestoAlerta).FuncionarioTienePermiso(int.Parse(HttpContext.Current.User.Identity.Name), permisoId) != 1)
            {
                Response.Redirect(oBase.carpetaRoot + "/PresupuestoAlerta/Default.aspx");
            }
            else
            {
                ((Label)Master.FindControl("TituloLbl")).Text = "Presupuesto ENAMI (Virtual Mode)";
                ValoresPorDefecto();
            }
        }
    }


    public void ValoresPorDefecto()
    {

        Base oBase = new Base();

        DataTable oCEBEList = new CEBEDAL(oBase.StrConexionPresupuestoAlerta).SelectAllDTByFK("", "", "", int.Parse(HttpContext.Current.User.Identity.Name));
        CebeDDL.DataSource = oCEBEList;
        CebeDDL.ValueField = "CEBECodigo";
        CebeDDL.TextField = "NombreCompuesto";
        CebeDDL.DataBind();
        CebeDDL.Items.Insert(0, new ListEditItem("Todos los CEBE...", "-1"));
        CebeDDL.Value = "-1";//"100235";


        DataTable oCECOList = new CECODAL(oBase.StrConexionPresupuestoAlerta).SelectAllDTByFK("", "", "", "", -1, int.Parse(HttpContext.Current.User.Identity.Name));
        CecoDDL.DataSource = oCECOList;
        CecoDDL.ValueField = "CECOCodigo";
        CecoDDL.TextField = "NombreCompuesto";
        CecoDDL.DataBind();
        CecoDDL.Items.Insert(0, new ListEditItem("Todos los CECO...", "-1"));
        CecoDDL.Value = "-1";//"100235";


        List<PresupuestoVersion> oPresupuestoVersionList = new PresupuestoVersionDAL(oBase.StrConexionPresupuestoAlerta).SelectAll();
        int anioMin = 0;
        int anioMax = 0;
        int anioActivo = 0;
        int versionActiva = 0;
        for (int i = 0; i < oPresupuestoVersionList.Count; i++)
        {
            if (oPresupuestoVersionList[i].Periodo < anioMin || anioMin == 0) anioMin = oPresupuestoVersionList[i].Periodo;
            if (oPresupuestoVersionList[i].Periodo > anioMax) anioMax = oPresupuestoVersionList[i].Periodo;

            if (oPresupuestoVersionList[i].Actual == 1 && anioActivo <= oPresupuestoVersionList[i].Periodo)
            {
                versionActiva = oPresupuestoVersionList[i].PresupuestoVersionId;
                anioActivo = oPresupuestoVersionList[i].Periodo;
            }
        }

        if (versionActiva <= 0 && oPresupuestoVersionList.Where(fil => fil.Periodo == anioActivo).ToList().Count > 0)
        {
            versionActiva = oPresupuestoVersionList.Where(fil => fil.Periodo == anioActivo).ToList()[0].PresupuestoVersionId;
        }
        VersionDDL.DataSource = oPresupuestoVersionList.Where(fil=> fil.Periodo == anioActivo).ToList();
        VersionDDL.ValueField = "PresupuestoVersionId";
        VersionDDL.TextField = "Nombre";
        VersionDDL.DataBind();
        VersionDDL.Value = versionActiva.ToString();
       
       
        PeriodoTxt.MinDate = DateTime.Parse("01-01-"+ anioMin.ToString());
        PeriodoTxt.MaxDate = DateTime.Parse("01-01-"+ anioMax.ToString());
        PeriodoTxt.Value = DateTime.Parse("01-01-" + anioActivo.ToString());


        Session["PeriodoSession"] = anioActivo;
        Session["VersionSession"] = versionActiva;        
        Session["CEBESession"] = CebeDDL.Value.ToString();
        Session["CECOSession"] = CecoDDL.Value.ToString();

        string usuario = HttpContext.Current.User.Identity.Name;
        int tienePermiso = new FuncionarioPermisoDAL(oBase.StrConexionPresupuestoAlerta).SelectAllByUsuario(usuario);

        PresupuestoTreeList_CustomCallback(null, null);

    }


    protected void PresupuestoTreeList_VirtualModeCreateChildren(object sender, TreeListVirtualModeCreateChildrenEventArgs e)
    {
        Base oBase = new Base();
        string codigo = e.NodeObject == null ? "" : e.NodeObject.ToString();
        string cebe = cebeSession != "-1" ? cebeSession : "";
        string ceco = cecoSession != "-1" ? cecoSession : "";
        int periodo = periodoSession;

        int origenIni = ceco != "" ? 1 : 0; 

        List<string> children = new List<string>();
        if (e.NodeObject == null)
        {
            DataTable oDataDT = new EstructuraPeriodoDAL(oBase.StrConexionPresupuestoAlerta).SelectAllDTByFKV2(origenIni, "", "", -1, periodo, cebe, ceco, "", "", int.Parse(HttpContext.Current.User.Identity.Name),1);
            foreach (DataRow row in oDataDT.Rows)
            {
                var codIni = row["CodigoIni"].ToString();
                children.Add(codIni);
            }           
        }
        else
        {
            DataTable oDataDT = new EstructuraPeriodoDAL(oBase.StrConexionPresupuestoAlerta).SelectAllDTByFKV2(-1, "", codigo, -1, periodo, "", "", "", "", int.Parse(HttpContext.Current.User.Identity.Name),1);
            foreach (DataRow row in oDataDT.Rows)
            {
                var codIni = row["CodigoIni"].ToString();
                children.Add(codIni);
            }
        }

        
        e.Children = children;

        PresupuestoTreeList.Toolbars[0].Items[0].Visible = false;
        PresupuestoTreeList.Toolbars[0].Items[1].Visible = false;
        PresupuestoTreeList.Toolbars[0].Items[2].Visible = false;

        List<PresupuestoVersionApertura> oPresupuestoVersionAperturaList = new PresupuestoVersionAperturaDAL(oBase.StrConexionPresupuestoAlerta).SelectAllByPresupuestoVersionId(versionSession);
        if (new FuncionarioPermisoDAL(oBase.StrConexionPresupuestoAlerta).FuncionarioTienePermiso(int.Parse(HttpContext.Current.User.Identity.Name), 199) == 1
        && oPresupuestoVersionAperturaList.Where(fil => fil.FechaInicio <= DateTime.Now && fil.FechaTermino >= DateTime.Now).ToList().Count > 0
        )
        {
            if (new FuncionarioPermisoDAL(oBase.StrConexionPresupuestoAlerta).FuncionarioTienePermiso(int.Parse(HttpContext.Current.User.Identity.Name), 1) == 1)
            {
                PresupuestoTreeList.Toolbars[0].Items[0].Visible = true;
            }
            PresupuestoTreeList.Toolbars[0].Items[1].Visible = true;
            PresupuestoTreeList.Toolbars[0].Items[2].Visible = true;
        }

       
    }

   
    protected void PresupuestoTreeList_VirtualModeNodeCreating(object sender, TreeListVirtualModeNodeCreatingEventArgs e)
    {
        Base oBase = new Base();
        string codigo = e.NodeObject.ToString();
        int periodo = periodoSession;

        e.NodeKeyValue = codigo;

        try
        {

            DataRow row = oDataTemp.Select("CodigoIni = '" + codigo + "'").FirstOrDefault();
            e.IsLeaf = row["EsUltimoNivel"].ToString() == "1" ? true : false;
            e.SetNodeValue("CodigoIni", row["CodigoIni"].ToString());
            e.SetNodeValue("Codigo", row["Codigo"].ToString());
            e.SetNodeValue("CuentaMadreCodigo", int.Parse(row["CuentaMadreCodigo"].ToString()));
            e.SetNodeValue("esAdmin", int.Parse(row["esAdmin"].ToString()));
            e.SetNodeValue("Nombre", row["Nombre"].ToString());
            e.SetNodeValue("EsUltimoNivel", row["EsUltimoNivel"].ToString());
            e.SetNodeValue("Origen", row["Origen"].ToString());
            //DESARROLLO DUHOVIT GRC

            double anioe = double.Parse(row["Anio_Total"].ToString());

            if (anioe == 0 || anioe == null)
            {
                e.SetNodeValue("Anio_Total", "-");
            }
            else
            {
                e.SetNodeValue("Anio_Total", anioe.ToString());
            }
            //DESARROLLO DUHOVIT GRC
            e.SetNodeValue("Total", double.Parse(row["Total"].ToString()));
            e.SetNodeValue("M1", double.Parse(row["M1"].ToString()));
            e.SetNodeValue("M2", double.Parse(row["M2"].ToString()));
            e.SetNodeValue("M3", double.Parse(row["M3"].ToString()));
            e.SetNodeValue("M4", double.Parse(row["M4"].ToString()));
            e.SetNodeValue("M5", double.Parse(row["M5"].ToString()));
            e.SetNodeValue("M6", double.Parse(row["M6"].ToString()));
            e.SetNodeValue("M7", double.Parse(row["M7"].ToString()));
            e.SetNodeValue("M8", double.Parse(row["M8"].ToString()));
            e.SetNodeValue("M9", double.Parse(row["M9"].ToString()));
            e.SetNodeValue("M10", double.Parse(row["M10"].ToString()));
            e.SetNodeValue("M11", double.Parse(row["M11"].ToString()));
            e.SetNodeValue("M12", double.Parse(row["M12"].ToString()));
            e.SetNodeValue("Adicional", double.Parse(row["Adicional"].ToString()));
            e.SetNodeValue("TieneSubNivel", int.Parse(row["TieneSubNivel"].ToString()));
            e.SetNodeValue("PuedePresupuestar", int.Parse(row["PuedePresupuestar"].ToString()));
            e.SetNodeValue("PuedeEliminar", int.Parse(row["PuedeEliminar"].ToString()));
        }
        catch {
            e.IsLeaf =  false;
            e.SetNodeValue("CodigoIni", codigo);
            e.SetNodeValue("Codigo", codigo);
            e.SetNodeValue("CuentaMadreCodigo",0);
            e.SetNodeValue("esAdmin", 0);
            e.SetNodeValue("Nombre", "---");
            e.SetNodeValue("EsUltimoNivel", 1);
            e.SetNodeValue("Origen", 0);
            e.SetNodeValue("Anio_Total", "-");
            e.SetNodeValue("Total", '-');
            e.SetNodeValue("M1", 0);
            e.SetNodeValue("M2", 0);
            e.SetNodeValue("M3", 0);
            e.SetNodeValue("M4", 0);
            e.SetNodeValue("M5", 0);
            e.SetNodeValue("M6", 0);
            e.SetNodeValue("M7", 0);
            e.SetNodeValue("M8", 0);
            e.SetNodeValue("M9", 0);
            e.SetNodeValue("M10", 0);
            e.SetNodeValue("M11", 0);
            e.SetNodeValue("M12", 0);
            e.SetNodeValue("Adicional", 0);
            e.SetNodeValue("TieneSubNivel", 0);
            e.SetNodeValue("PuedePresupuestar", 0);
            e.SetNodeValue("PuedeEliminar", 0);
        }
    }



    protected void PresupuestoTreeList_CustomCallback(object sender, TreeListCustomCallbackEventArgs e)
    {
        Base oBase = new Base();

        versionSession = int.Parse(VersionDDL.Value.ToString());
        periodoSession = int.Parse(PeriodoTxt.Text.ToString());
        cebeSession = CebeDDL.Value.ToString();
        cecoSession = CecoDDL.Value.ToString();

        if (oDataTemp != null) oDataTemp.Rows.Clear();
        oDataTemp = new EstructuraPeriodoDAL(oBase.StrConexionPresupuestoAlerta).SelectAllDTByFKV2(-1, "", "", versionSession, periodoSession, "", "", "", "", int.Parse(HttpContext.Current.User.Identity.Name), 0);

        PresupuestoTreeList.RefreshVirtualTree();
    }

    protected void VersionDDL_Callback(object sender, CallbackEventArgsBase e)
    {
        Base oBase = new Base();
     
        List<PresupuestoVersion> oPresupuestoVersionList = new PresupuestoVersionDAL(oBase.StrConexionPresupuestoAlerta).SelectAll().Where(fil => fil.Periodo == int.Parse(PeriodoTxt.Text.ToString())).ToList();
        VersionDDL.DataSource = oPresupuestoVersionList;
        VersionDDL.ValueField = "PresupuestoVersionId";
        VersionDDL.TextField = "Nombre";
        VersionDDL.DataBind();

        VersionDDL.Value = oPresupuestoVersionList[0].PresupuestoVersionId.ToString();
        versionSession = int.Parse(VersionDDL.Value.ToString());
        periodoSession = int.Parse(PeriodoTxt.Text.ToString());

    }

    protected void AgregarCebeBtn_Click(object sender, EventArgs e)
    {
        Base oBase = new Base();
        Utility.BaseTransaccion SqlClientUtility = new Utility.BaseTransaccion();

        EstructuraPeriodoDAL oEstructuraPeriodoDAL = new EstructuraPeriodoDAL(oBase.StrConexionPresupuestoAlerta);

        SqlClientUtility.IniciarTransaccion(oBase.StrConexionPresupuestoAlerta);//Inicio la transaccion
        try
        {
            int anio = int.Parse(PeriodoAddHid.Value);       
            var codigos = CebesAddHid.Value.Split(',');

            for (int i = 0; i < codigos.Length; i++)
            {
                var codigoAdd = codigos[i];

                if (codigoAdd != "")
                {
                    EstructuraPeriodo oEstructuraPeriodo = new EstructuraPeriodo();
                    oEstructuraPeriodo.Periodo = anio;
                    oEstructuraPeriodo.CodigoIni = codigoAdd;
                    oEstructuraPeriodo.Origen = 0;
                    oEstructuraPeriodo.Codigo = codigoAdd;
                    oEstructuraPeriodo.CodigoPadre = "";
                    oEstructuraPeriodo.CEBECodigo = codigoAdd;
                    oEstructuraPeriodo.CECOCodigo = "";
                    oEstructuraPeriodo.CuentaSAPCodigo = "";
                    oEstructuraPeriodo.ArticuloGrupoCodigo = "";
                    oEstructuraPeriodo.MaterialCodigo = "";
                    oEstructuraPeriodo.EsUltimoNivel = 0;

                    oEstructuraPeriodoDAL.Insert(oEstructuraPeriodo, SqlClientUtility);
                }
            }

            SqlClientUtility.ProcesarTransaccion();
            try
            {
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "presupuestotreelist.PerformCallback();alertify.success('CEBE(s) agregado(s)');", true);
            }
            catch { }
        }
        catch
        {
            SqlClientUtility.DeshacerTransaccion();  /// si falla la transaccion
            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "alertify.error('Error al Grabar, Consulte al Administrador');", true);
        }
    }

    protected void AgregarCecoBtn_Click(object sender, EventArgs e)
    {
        Base oBase = new Base();
        Utility.BaseTransaccion SqlClientUtility = new Utility.BaseTransaccion();

        EstructuraPeriodoDAL oEstructuraPeriodoDAL = new EstructuraPeriodoDAL(oBase.StrConexionPresupuestoAlerta);

        SqlClientUtility.IniciarTransaccion(oBase.StrConexionPresupuestoAlerta);//Inicio la transaccion
        try
        {
            int anio = int.Parse(PeriodoAddHid.Value);
            string cebeCodigo = CebeSelHid.Value;
            var codigos = CecosAddHid.Value.Split(',');

            for (int i = 0; i < codigos.Length; i++)
            {
                var codigoAdd = codigos[i];

                if (codigoAdd != "")
                {
                    EstructuraPeriodo oEstructuraPeriodo = new EstructuraPeriodo();
                    oEstructuraPeriodo.Periodo = anio;
                    oEstructuraPeriodo.CodigoIni = cebeCodigo+";"+codigoAdd;
                    oEstructuraPeriodo.Origen = 1;
                    oEstructuraPeriodo.Codigo = codigoAdd;
                    oEstructuraPeriodo.CodigoPadre = cebeCodigo;
                    oEstructuraPeriodo.CEBECodigo = cebeCodigo;
                    oEstructuraPeriodo.CECOCodigo = codigoAdd;
                    oEstructuraPeriodo.CuentaSAPCodigo = "";
                    oEstructuraPeriodo.ArticuloGrupoCodigo = "";
                    oEstructuraPeriodo.MaterialCodigo = "";
                    oEstructuraPeriodo.EsUltimoNivel = 0;

                    oEstructuraPeriodoDAL.Insert(oEstructuraPeriodo, SqlClientUtility);
                }
            }

            SqlClientUtility.ProcesarTransaccion();
            try
            {
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "presupuestotreelist.PerformCallback();alertify.success('CECO(s) agregado(s)');", true);
            }
            catch { }
        }
        catch
        {
            SqlClientUtility.DeshacerTransaccion();  /// si falla la transaccion
            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "alertify.error('Error al Grabar, Consulte al Administrador');", true);
        }
    }

    protected void AgregarCSapBtn_Click(object sender, EventArgs e)
    {
        Base oBase = new Base();
        Utility.BaseTransaccion SqlClientUtility = new Utility.BaseTransaccion();

        EstructuraPeriodoDAL oEstructuraPeriodoDAL = new EstructuraPeriodoDAL(oBase.StrConexionPresupuestoAlerta);

        SqlClientUtility.IniciarTransaccion(oBase.StrConexionPresupuestoAlerta);//Inicio la transaccion
        try
        {
            int anio = int.Parse(PeriodoAddHid.Value);
            string cebeCodigo = CebeSelHid.Value;
            string cecoCodigo = CecoSelHis.Value;
            var codigos = CuentasSAPAddHid.Value.Split(',');

            for (int i = 0; i < codigos.Length; i++)
            {
                var codigoAdd = codigos[i];

                if (codigoAdd != "")
                {
                    EstructuraPeriodo oEstructuraPeriodo = new EstructuraPeriodo();
                    oEstructuraPeriodo.Periodo = anio;
                    oEstructuraPeriodo.CodigoIni = cebeCodigo + ";" + cecoCodigo + ";" + codigoAdd;
                    oEstructuraPeriodo.Origen = 2;
                    oEstructuraPeriodo.Codigo = codigoAdd;
                    oEstructuraPeriodo.CodigoPadre = cebeCodigo + ";" + cecoCodigo;
                    oEstructuraPeriodo.CEBECodigo = cebeCodigo;
                    oEstructuraPeriodo.CECOCodigo = cecoCodigo;
                    oEstructuraPeriodo.CuentaSAPCodigo = codigoAdd;
                    oEstructuraPeriodo.ArticuloGrupoCodigo = "";
                    oEstructuraPeriodo.MaterialCodigo = "";
                    oEstructuraPeriodo.EsUltimoNivel = codigoAdd != "61220013" && codigoAdd != "61220018" ? 1 : 0;

                    oEstructuraPeriodoDAL.Insert(oEstructuraPeriodo, SqlClientUtility);
                }
            }

            SqlClientUtility.ProcesarTransaccion();
            try
            {
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "presupuestotreelist.PerformCallback();alertify.success('Cuenta(s) SAP agregada(s)');", true);
            }
            catch { }
        }
        catch
        {
            SqlClientUtility.DeshacerTransaccion();  /// si falla la transaccion
            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "alertify.error('Error al Grabar, Consulte al Administrador');", true);
        }
    }

    protected void AgregarGArticuloBtn_Click(object sender, EventArgs e)
    {
        Base oBase = new Base();
        Utility.BaseTransaccion SqlClientUtility = new Utility.BaseTransaccion();

        EstructuraPeriodoDAL oEstructuraPeriodoDAL = new EstructuraPeriodoDAL(oBase.StrConexionPresupuestoAlerta);

        SqlClientUtility.IniciarTransaccion(oBase.StrConexionPresupuestoAlerta);//Inicio la transaccion

        try
        {
            int anio = int.Parse(PeriodoAddHid.Value);
            string cebeCodigo = CebeSelHid.Value;
            string cecoCodigo = CecoSelHis.Value;
            string csapCodigo = CSapSelHid.Value;
            var codigos = ArticulosGrupoAddHid.Value.Split(',');

            for (int i = 0; i < codigos.Length; i++)
            {
                var codigoAdd = codigos[i];

                if (codigoAdd != "")
                {
                    EstructuraPeriodo oEstructuraPeriodo = new EstructuraPeriodo();
                    oEstructuraPeriodo.Periodo = anio;
                    oEstructuraPeriodo.CodigoIni = cebeCodigo + ";" + cecoCodigo + ";"+ csapCodigo + ";" + codigoAdd;
                    oEstructuraPeriodo.Origen = 3;
                    oEstructuraPeriodo.Codigo = codigoAdd;
                    oEstructuraPeriodo.CodigoPadre = cebeCodigo + ";" + cecoCodigo + ";" + csapCodigo;
                    oEstructuraPeriodo.CEBECodigo = cebeCodigo;
                    oEstructuraPeriodo.CECOCodigo = cecoCodigo;
                    oEstructuraPeriodo.CuentaSAPCodigo = csapCodigo;
                    oEstructuraPeriodo.ArticuloGrupoCodigo = codigoAdd;
                    oEstructuraPeriodo.MaterialCodigo = "";
                    oEstructuraPeriodo.EsUltimoNivel = codigoAdd.Substring(0,1) == "M" || (codigoAdd.Substring(0, 1) == "S" && anio>= 2024) ? 0 : 1;

                    oEstructuraPeriodoDAL.Insert(oEstructuraPeriodo, SqlClientUtility);
                }
            }

            SqlClientUtility.ProcesarTransaccion();
            try
            {
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "presupuestotreelist.PerformCallback();alertify.success('Cuenta(s) SAP agregada(s)');", true);
            }
            catch { }
        }
        catch
        {
            SqlClientUtility.DeshacerTransaccion();  /// si falla la transaccion
            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "alertify.error('Error al Grabar, Consulte al Administrador');", true);
        }
    }

    protected void AgregarMaterialBtn_Click(object sender, EventArgs e)
    {
        Base oBase = new Base();
        Utility.BaseTransaccion SqlClientUtility = new Utility.BaseTransaccion();

        EstructuraPeriodoDAL oEstructuraPeriodoDAL = new EstructuraPeriodoDAL(oBase.StrConexionPresupuestoAlerta);

        SqlClientUtility.IniciarTransaccion(oBase.StrConexionPresupuestoAlerta);//Inicio la transaccion
        try
        {
            int anio = int.Parse(PeriodoAddHid.Value);
            string cebeCodigo = CebeSelHid.Value;
            string cecoCodigo = CecoSelHis.Value;
            string csapCodigo = CSapSelHid.Value;
            string agrupoCodigo = AGrupoSelHid.Value;
            var codigos = MaterialesAddHid.Value.Split(',');

            for (int i = 0; i < codigos.Length; i++)
            {
                var codigoAdd = codigos[i];

                if (codigoAdd != "")
                {
                    EstructuraPeriodo oEstructuraPeriodo = new EstructuraPeriodo();
                    oEstructuraPeriodo.Periodo = anio;
                    oEstructuraPeriodo.CodigoIni = cebeCodigo + ";" + cecoCodigo + ";" + csapCodigo + ";" + agrupoCodigo + ";" + codigoAdd;
                    oEstructuraPeriodo.Origen = 4;
                    oEstructuraPeriodo.Codigo = codigoAdd;
                    oEstructuraPeriodo.CodigoPadre = cebeCodigo + ";" + cecoCodigo + ";" + csapCodigo + ";" + agrupoCodigo;
                    oEstructuraPeriodo.CEBECodigo = cebeCodigo;
                    oEstructuraPeriodo.CECOCodigo = cecoCodigo;
                    oEstructuraPeriodo.CuentaSAPCodigo = csapCodigo;
                    oEstructuraPeriodo.ArticuloGrupoCodigo = agrupoCodigo;
                    oEstructuraPeriodo.MaterialCodigo = codigoAdd;
                    oEstructuraPeriodo.EsUltimoNivel = 1;

                    oEstructuraPeriodoDAL.Insert(oEstructuraPeriodo, SqlClientUtility);
                }
            }

            SqlClientUtility.ProcesarTransaccion();
            try
            {
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "presupuestotreelist.PerformCallback();alertify.success('Material(es)/Contrato(s) agregado(s)');", true);
            }
            catch { }
        }
        catch
        {
            SqlClientUtility.DeshacerTransaccion();  /// si falla la transaccion
            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "alertify.error('Error al Grabar, Consulte al Administrador');", true);
        }
    }

    protected void EliminarItemBtn_Click(object sender, EventArgs e)
    {
        Base oBase = new Base();
        Utility.BaseTransaccion SqlClientUtility = new Utility.BaseTransaccion();

		string codigoIni = CodigoIniEliminarHid.Value;
		int anio = int.Parse(PeriodoAddHid.Value);
		List<PresupuestoVersion> oPresupuestoVersionList = new PresupuestoVersionDAL(oBase.StrConexionPresupuestoAlerta).SelectAll().Where(fil => fil.Periodo == anio).ToList();
        if (oPresupuestoVersionList.Count > 1)
        {
			ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "alertify.error('No puede eliminar conceptos ligados a otras versiones');", true);
		}
		else
        {
            EstructuraPeriodoDAL oEstructuraPeriodoDAL = new EstructuraPeriodoDAL(oBase.StrConexionPresupuestoAlerta);
            PresupuestoVersionHistorialDAL oPresupuestoVersionHistorialDAL = new PresupuestoVersionHistorialDAL(oBase.StrConexionPresupuestoAlerta);

            SqlClientUtility.IniciarTransaccion(oBase.StrConexionPresupuestoAlerta);//Inicio la transaccion
            try
            {



                oEstructuraPeriodoDAL.Delete(anio, codigoIni, SqlClientUtility);

                var ids = codigoIni.Split(';');

                string cebeCodigo = ids[0];
                string cecoCodigo = ids[1];
                string cuentaSAPCodigo = ids[2];

                string articuloGrupoCodigo = ids.Length > 3 ? ids[3] : "";
                string materialCodigo = ids.Length > 4 ? ids[4] : "";

                PresupuestoVersionHistorial oPresupuestoVersionHistorial = new PresupuestoVersionHistorial();
                oPresupuestoVersionHistorial.PresupuestoVersionId = versionSession;
                oPresupuestoVersionHistorial.CEBECodigo = cebeCodigo;
                oPresupuestoVersionHistorial.CECOCodigo = cecoCodigo;
                oPresupuestoVersionHistorial.CuentaSAPCodigo = cuentaSAPCodigo;
                oPresupuestoVersionHistorial.ArticuloGrupoCodigo = articuloGrupoCodigo;
                oPresupuestoVersionHistorial.MaterialCodigo = materialCodigo;
                oPresupuestoVersionHistorial.Tipo = 0;
                oPresupuestoVersionHistorial.Origen = 1;
                oPresupuestoVersionHistorial.OrigenStr = "Concepto Eliminado";
                oPresupuestoVersionHistorial.Obsevacion = "";
                oPresupuestoVersionHistorial.UsuarioAdd = int.Parse(HttpContext.Current.User.Identity.Name);
                oPresupuestoVersionHistorial.FechaAdd = DateTime.Now;
                oPresupuestoVersionHistorialDAL.Insert(oPresupuestoVersionHistorial, SqlClientUtility);

                SqlClientUtility.ProcesarTransaccion();
                try
                {
                    ActualizarRaiz();
                }
                catch { }
            }
            catch
            {
                SqlClientUtility.DeshacerTransaccion();  /// si falla la transaccion
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "alertify.error('Error al Eliminar(1), Consulte al Administrador');", true);
            }
        }
    }



    public void ActualizarRaiz()
    {
        Base oBase = new Base();
        Utility.BaseTransaccion SqlClientUtility = new Utility.BaseTransaccion();

        PresupuestoVersionDataDAL oPresupuestoVersionDataDAL = new PresupuestoVersionDataDAL(oBase.StrConexionPresupuestoAlerta);

        SqlClientUtility.IniciarTransaccion(oBase.StrConexionPresupuestoAlerta);//Inicio la transaccion
        try
        {
            int origen = int.Parse(OrigenEliminarHid.Value);
            int version = int.Parse(VersionEliminarHid.Value);
            string codigo = CodigoIniEliminarHid.Value;

            var ids = codigo.Split(';');

            //string cebeCodigo = ids[0];
            //string cecoCodigo = origen >= 2 ?  ids[1] : "";
            //string cuentaSAPCodigo = origen >= 3 ? ids[2] : "";

            //string articuloGrupoCodigo = origen == 4 || origen == 5 ? ids[3] : "";
            //string materialCodigo = origen == 5 ? ids[3] : "";

            string cebeCodigo = ids[0];
            string cecoCodigo = ids[1];
            string cuentaSAPCodigo = ids[2];

            string articuloGrupoCodigo = ids.Length > 3 ? ids[3] : "";
            string materialCodigo = ids.Length > 4 ? ids[4] : "";

            oPresupuestoVersionDataDAL.UpdateRaiz(version, cebeCodigo, cecoCodigo, cuentaSAPCodigo, articuloGrupoCodigo, materialCodigo, SqlClientUtility);

            SqlClientUtility.ProcesarTransaccion();
            try
            {
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "presupuestotreelist.PerformCallback();alertify.success('Eliminado correctamente');", true);
            }
            catch { }
        }
        catch
        {
            SqlClientUtility.DeshacerTransaccion();  /// si falla la transaccion
            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "alertify.error('Error al Eliminar(2), Consulte al Administrador');", true);
        }
    }


    protected void EliminarPresupuestoBtn_Click(object sender, EventArgs e)
    {
        Base oBase = new Base();
        Utility.BaseTransaccion SqlClientUtility = new Utility.BaseTransaccion();

        PresupuestoVersionDataDAL oPresupuestoVersionDataDAL = new PresupuestoVersionDataDAL(oBase.StrConexionPresupuestoAlerta);

        SqlClientUtility.IniciarTransaccion(oBase.StrConexionPresupuestoAlerta);
        try
        {
            oPresupuestoVersionDataDAL.DeleteAllByVersionPermiso(int.Parse(HttpContext.Current.User.Identity.Name), int.Parse(VersionDDL.Value.ToString()), SqlClientUtility);

            SqlClientUtility.ProcesarTransaccion();
            try
            {
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "window.parent.Actualizar();\nwindow.parent.Shadowbox.close();", true);
            }
            catch { }
        }
        catch
        {
            SqlClientUtility.DeshacerTransaccion();  /// si falla la transaccion
            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "script", "alertify.error('Error al Grabar, Consulte al Administrador');", true);
        }
    }

    protected void PresupuestoTreeList_ToolbarItemClick(object source, TreeListToolbarItemClickEventArgs e)
    {
        switch (e.Item.Name)
        {

            case "ExportToXLSXH":
                ExportarInfoHorizontal();
                break;
            case "ExportToXLSXV":
                ExportarInfoVertical();
                break;
            case "ExportToXLSXPC":
                ExportarPlantillaCarga();
                break;          
            default:
                break;
        }
    }

    public void ExportarInfoHorizontal()
    {
        Base oBase = new Base();

        int version = int.Parse(VersionDDL.Value.ToString());

        Color rowColor = System.Drawing.ColorTranslator.FromHtml("#167b67");
        Color rowColorF = System.Drawing.ColorTranslator.FromHtml("#ffffff");

        DateTime fechaMin = DateTime.Parse("01-01-1900");

        string usuario = HttpContext.Current.User.Identity.Name;
        int tienePermiso = new FuncionarioPermisoDAL(oBase.StrConexionPresupuestoAlerta).SelectAllByUsuario(usuario);

        if (tienePermiso != 1)
        {
            using (Workbook workbook = new Workbook())
            {

                workbook.Unit = DevExpress.Office.DocumentUnit.Point;

                #region Movimientos
                DataTable oData = new PresupuestoVersionDataDAL(oBase.StrConexionPresupuestoAlerta).ExportarByFK(int.Parse(HttpContext.Current.User.Identity.Name), version, 0);
                Worksheet worksheetMovimiento = workbook.Worksheets[0];
                worksheetMovimiento.Name = "Plantilla Carga";
                workbook.BeginUpdate();
                int periodo = periodoSession;

                try
                {
                    worksheetMovimiento.OutlineOptions.SummaryRowsBelow = false;

                    //Cabecera
                    worksheetMovimiento.Cells["A1"].Value = "CEBE";
                    worksheetMovimiento.Cells["B1"].Value = "CEBE NOMBRE";
                    worksheetMovimiento.Cells["C1"].Value = "CECO";
                    worksheetMovimiento.Cells["D1"].Value = "CECO NOMBRE";
                    worksheetMovimiento.Cells["E1"].Value = "CUENTA";
                    worksheetMovimiento.Cells["F1"].Value = "CUENTA NOMBRE";
                    worksheetMovimiento.Cells["G1"].Value = "GRUPO DE ARTICULO";
                    worksheetMovimiento.Cells["H1"].Value = "GRUPO DE ARTICULO NOMBRE";
                    worksheetMovimiento.Cells["I1"].Value = "CODIGO DE MATERIAL";
                    worksheetMovimiento.Cells["J1"].Value = "CODIGO DE MATERIAL NOMBRE";
                    worksheetMovimiento.Cells["K1"].Value = "AÑO HASTA";

                    worksheetMovimiento.Cells["L1"].Value = "ENERO";
                    worksheetMovimiento.Cells["M1"].Value = "FEBRERO";
                    worksheetMovimiento.Cells["N1"].Value = "MARZO";
                    worksheetMovimiento.Cells["O1"].Value = "ABRIL";
                    worksheetMovimiento.Cells["P1"].Value = "MAYO";
                    worksheetMovimiento.Cells["Q1"].Value = "JUNIO";
                    worksheetMovimiento.Cells["R1"].Value = "JULIO";
                    worksheetMovimiento.Cells["S1"].Value = "AGOSTO";
                    worksheetMovimiento.Cells["T1"].Value = "SEPTIEMBRE";
                    worksheetMovimiento.Cells["U1"].Value = "OCTUBRE";
                    worksheetMovimiento.Cells["V1"].Value = "NOVIEMBRE";
                    worksheetMovimiento.Cells["W1"].Value = "DICIEMBRE";
                    worksheetMovimiento.Cells["X1"].Value = "TOTAL";

                    worksheetMovimiento.Range["A1:K1"].ColumnWidth = 75;
                    worksheetMovimiento.Cells["B1"].ColumnWidth = 170;
                    worksheetMovimiento.Cells["D1"].ColumnWidth = 170;
                    worksheetMovimiento.Cells["F1"].ColumnWidth = 170;
                    worksheetMovimiento.Cells["H1"].ColumnWidth = 170;
                    worksheetMovimiento.Cells["J1"].ColumnWidth = 170;

                    worksheetMovimiento.Range["L1:X1"].ColumnWidth = 70;
                    worksheetMovimiento.Range["A1:X1"].Alignment.WrapText = true;
                    worksheetMovimiento.Range["A1:X1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                    worksheetMovimiento.Range["A1:X1"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                    worksheetMovimiento.Range["A1:J1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                    worksheetMovimiento.Range["A1:X1"].Font.Bold = true;
                    worksheetMovimiento.Range["A1:X1"].Fill.BackgroundColor = rowColor;
                    worksheetMovimiento.Range["A1:X1"].Font.Color = rowColorF;

                    //Detalle                
                    int rowActual = 1;
                    foreach (DataRow row in oData.Rows)
                    {

                        worksheetMovimiento.Columns["A"][rowActual].Value = int.Parse(row["CEBECodigo"].ToString());
                        worksheetMovimiento.Columns["B"][rowActual].Value = row["CEBEN"].ToString();
                        worksheetMovimiento.Columns["C"][rowActual].Value = row["CECOCodigo"].ToString();
                        worksheetMovimiento.Columns["D"][rowActual].Value = row["CECON"].ToString();
                        worksheetMovimiento.Columns["E"][rowActual].Value = int.Parse(row["CuentaSAPCodigo"].ToString());
                        worksheetMovimiento.Columns["F"][rowActual].Value = row["CuentaSAPN"].ToString();
                        worksheetMovimiento.Columns["G"][rowActual].Value = row["ArticuloGrupoCodigo"].ToString();
                        worksheetMovimiento.Columns["H"][rowActual].Value = row["ArticuloGrupoN"].ToString();
                        worksheetMovimiento.Columns["I"][rowActual].Value = row["MaterialCodigo"].ToString();
                        worksheetMovimiento.Columns["J"][rowActual].Value = row["MaterialN"].ToString();
                        worksheetMovimiento.Columns["K"][rowActual].Value = row["Anio_Total"].ToString();

                        double m1 = 0;
                        double.TryParse(row["M1"].ToString(), out m1);
                        double m2 = 0;
                        double.TryParse(row["M2"].ToString(), out m2);
                        double m3 = 0;
                        double.TryParse(row["M3"].ToString(), out m3);
                        double m4 = 0;
                        double.TryParse(row["M4"].ToString(), out m4);
                        double m5 = 0;
                        double.TryParse(row["M5"].ToString(), out m5);
                        double m6 = 0;
                        double.TryParse(row["M6"].ToString(), out m6);
                        double m7 = 0;
                        double.TryParse(row["M7"].ToString(), out m7);
                        double m8 = 0;
                        double.TryParse(row["M8"].ToString(), out m8);
                        double m9 = 0;
                        double.TryParse(row["M9"].ToString(), out m9);
                        double m10 = 0;
                        double.TryParse(row["M10"].ToString(), out m10);
                        double m11 = 0;
                        double.TryParse(row["M11"].ToString(), out m11);
                        double m12 = 0;
                        double.TryParse(row["M12"].ToString(), out m12);

                        double total = 0;
                        double.TryParse(row["Total"].ToString(), out total);

                        worksheetMovimiento.Columns["L"][rowActual].Value = m1;
                        worksheetMovimiento.Columns["M"][rowActual].Value = m2;
                        worksheetMovimiento.Columns["N"][rowActual].Value = m3;
                        worksheetMovimiento.Columns["O"][rowActual].Value = m4;
                        worksheetMovimiento.Columns["P"][rowActual].Value = m5;
                        worksheetMovimiento.Columns["Q"][rowActual].Value = m6;
                        worksheetMovimiento.Columns["R"][rowActual].Value = m7;
                        worksheetMovimiento.Columns["S"][rowActual].Value = m8;
                        worksheetMovimiento.Columns["T"][rowActual].Value = m9;
                        worksheetMovimiento.Columns["U"][rowActual].Value = m10;
                        worksheetMovimiento.Columns["V"][rowActual].Value = m11;
                        worksheetMovimiento.Columns["W"][rowActual].Value = m12;

                        worksheetMovimiento.Columns["X"][rowActual].Value = total;


                        rowActual++;
                    }

                    int lastRow = rowActual + 1; 

                    string finalMergeRange = String.Format("L{0}:X{0}", lastRow);
                    worksheetMovimiento.Range[finalMergeRange].Merge();
                    worksheetMovimiento.Cells[String.Format("L{0}", lastRow)].Value = periodo.ToString();

                    worksheetMovimiento.Range[finalMergeRange].Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center;
                    worksheetMovimiento.Range[finalMergeRange].Alignment.Vertical = DevExpress.Spreadsheet.SpreadsheetVerticalAlignment.Center;

                    worksheetMovimiento.Cells[String.Format("L{0}", lastRow)].NumberFormat = "@";

                    worksheetMovimiento.Range[finalMergeRange].Font.Bold = true;
                    worksheetMovimiento.Range[finalMergeRange].Fill.BackgroundColor = rowColor;
                    worksheetMovimiento.Range[finalMergeRange].Font.Color = rowColorF;


                    worksheetMovimiento.Range["L:X"].NumberFormat = "#,##0.00";

                    worksheetMovimiento.Range["A1:X" + (rowActual)].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);



                }
                finally
                {
                    workbook.EndUpdate();
                }

                #endregion


                string nombreExcel = VersionDDL.Text + " - " + PeriodoTxt.Text + " (Horizontal).xlsx";
                //// Save the document file under the specified name.
                workbook.SaveDocument(Server.MapPath(nombreExcel), DocumentFormat.OpenXml);

                FileInfo myfile = new FileInfo(Server.MapPath(nombreExcel));
                // Clear the content of the response
                byte[] archivo;
                Response.ClearContent();
                Response.Clear();
                // Set the ContentType
                Response.Cache.SetCacheability(HttpCacheability.Private);
                Response.ContentType = "application/vnd.ms-excel";//application/vnd.ms-excel   //application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
                                                                  // Add the file name and attachment, which will force the open/cancel/save dialog box to show, to the header
                Response.AddHeader("Content-Disposition", "attachment; filename=" + nombreExcel);
                // Add the file size into the response header
                Response.AddHeader("Content-Length", (myfile.Length).ToString());
                // Write the file into the response (TransmitFile is for ASP.NET 2.0. In ASP.NET 1.1 you have to use WriteFile instead)
                archivo = File.ReadAllBytes(myfile.FullName);
                Response.BinaryWrite(archivo);

                //Response.Flush();
                File.Delete(Server.MapPath(nombreExcel));


                //Response.WriteFile(myfile.FullName);
                // End the response
                Response.End();


                //// Export the document to PDF.
                //workbook.ExportToPdf("TestDoc.pdf");
            }
        }
        else
        {
            // Create a new workbook.
            using (Workbook workbook = new Workbook())
            {

                workbook.Unit = DevExpress.Office.DocumentUnit.Point;


                DataTable oData = new PresupuestoVersionDataDAL(oBase.StrConexionPresupuestoAlerta).ExportarByFK(int.Parse(HttpContext.Current.User.Identity.Name), version, 0);
                Worksheet worksheetMovimiento = workbook.Worksheets[0];
                worksheetMovimiento.Name = "Plantilla Carga";
                workbook.BeginUpdate();
                int periodo = periodoSession;
                try
                {
                    worksheetMovimiento.OutlineOptions.SummaryRowsBelow = false;

                    int maxAnioTotal = 0;

                    foreach (DataRow row in oData.Rows)
                    {
                        string anioTotalStr = row["Anio_Total"].ToString();
                        int anioTotal = 0;


                        if (!string.IsNullOrEmpty(anioTotalStr) && anioTotalStr != "-" && anioTotalStr != "0")
                        {
                            int.TryParse(anioTotalStr, out anioTotal);
                        }

                        // Actualizar el valor máximo
                        if (anioTotal > maxAnioTotal)
                        {
                            maxAnioTotal = anioTotal;
                        }

                    }

                    if (maxAnioTotal - periodo == 2)
                    {
                        //Cabecera
                        worksheetMovimiento.Cells["A1"].Value = "CEBE";
                        worksheetMovimiento.Cells["B1"].Value = "CEBE NOMBRE";
                        worksheetMovimiento.Cells["C1"].Value = "CECO";
                        worksheetMovimiento.Cells["D1"].Value = "CECO NOMBRE";
                        worksheetMovimiento.Cells["E1"].Value = "CUENTA";
                        worksheetMovimiento.Cells["F1"].Value = "CUENTA NOMBRE";
                        worksheetMovimiento.Cells["G1"].Value = "GRUPO DE ARTICULO";
                        worksheetMovimiento.Cells["H1"].Value = "GRUPO DE ARTICULO NOMBRE";
                        worksheetMovimiento.Cells["I1"].Value = "CODIGO DE MATERIAL";
                        worksheetMovimiento.Cells["J1"].Value = "CODIGO DE MATERIAL NOMBRE";
                        worksheetMovimiento.Cells["K1"].Value = "AÑO HASTA";

                        worksheetMovimiento.Cells["L1"].Value = "ENERO";
                        worksheetMovimiento.Cells["M1"].Value = "FEBRERO";
                        worksheetMovimiento.Cells["N1"].Value = "MARZO";
                        worksheetMovimiento.Cells["O1"].Value = "ABRIL";
                        worksheetMovimiento.Cells["P1"].Value = "MAYO";
                        worksheetMovimiento.Cells["Q1"].Value = "JUNIO";
                        worksheetMovimiento.Cells["R1"].Value = "JULIO";
                        worksheetMovimiento.Cells["S1"].Value = "AGOSTO";
                        worksheetMovimiento.Cells["T1"].Value = "SEPTIEMBRE";
                        worksheetMovimiento.Cells["U1"].Value = "OCTUBRE";
                        worksheetMovimiento.Cells["V1"].Value = "NOVIEMBRE";
                        worksheetMovimiento.Cells["W1"].Value = "DICIEMBRE";
                        worksheetMovimiento.Cells["X1"].Value = "TOTAL";

                        worksheetMovimiento.Cells["Y1"].Value = "ENERO";
                        worksheetMovimiento.Cells["Z1"].Value = "FEBRERO";
                        worksheetMovimiento.Cells["AA1"].Value = "MARZO";
                        worksheetMovimiento.Cells["AB1"].Value = "ABRIL";
                        worksheetMovimiento.Cells["AC1"].Value = "MAYO";
                        worksheetMovimiento.Cells["AD1"].Value = "JUNIO";
                        worksheetMovimiento.Cells["AE1"].Value = "JULIO";
                        worksheetMovimiento.Cells["AF1"].Value = "AGOSTO";
                        worksheetMovimiento.Cells["AG1"].Value = "SEPTIEMBRE";
                        worksheetMovimiento.Cells["AH1"].Value = "OCTUBRE";
                        worksheetMovimiento.Cells["AI1"].Value = "NOVIEMBRE";
                        worksheetMovimiento.Cells["AJ1"].Value = "DICIEMBRE";
                        worksheetMovimiento.Cells["AK1"].Value = "TOTAL";

                        worksheetMovimiento.Cells["AL1"].Value = "ENERO";
                        worksheetMovimiento.Cells["AM1"].Value = "FEBRERO";
                        worksheetMovimiento.Cells["AN1"].Value = "MARZO";
                        worksheetMovimiento.Cells["AO1"].Value = "ABRIL";
                        worksheetMovimiento.Cells["AP1"].Value = "MAYO";
                        worksheetMovimiento.Cells["AQ1"].Value = "JUNIO";
                        worksheetMovimiento.Cells["AR1"].Value = "JULIO";
                        worksheetMovimiento.Cells["AS1"].Value = "AGOSTO";
                        worksheetMovimiento.Cells["AT1"].Value = "SEPTIEMBRE";
                        worksheetMovimiento.Cells["AU1"].Value = "OCTUBRE";
                        worksheetMovimiento.Cells["AV1"].Value = "NOVIEMBRE";
                        worksheetMovimiento.Cells["AW1"].Value = "DICIEMBRE";
                        worksheetMovimiento.Cells["AX1"].Value = "TOTAL";


                        worksheetMovimiento.Range["A1:K1"].ColumnWidth = 75;
                        worksheetMovimiento.Cells["B1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["D1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["F1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["H1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["J1"].ColumnWidth = 170;

                        worksheetMovimiento.Range["L1:AX1"].ColumnWidth = 90;
                        worksheetMovimiento.Range["A1:AX1"].Alignment.WrapText = true;
                        worksheetMovimiento.Range["A1:AX1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                        worksheetMovimiento.Range["A1:AX1"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                        worksheetMovimiento.Range["A1:J1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                        worksheetMovimiento.Range["A1:AX1"].Font.Bold = true;
                        worksheetMovimiento.Range["A1:AX1"].Fill.BackgroundColor = rowColor;
                        worksheetMovimiento.Range["A1:AX1"].Font.Color = rowColorF;

                        //Detalle                
                        int rowActual = 1;
                        foreach (DataRow row in oData.Rows)
                        {

                            worksheetMovimiento.Columns["A"][rowActual].Value = int.Parse(row["CEBECodigo"].ToString());
                            worksheetMovimiento.Columns["B"][rowActual].Value = row["CEBEN"].ToString();
                            worksheetMovimiento.Columns["C"][rowActual].Value = row["CECOCodigo"].ToString();
                            worksheetMovimiento.Columns["D"][rowActual].Value = row["CECON"].ToString();
                            worksheetMovimiento.Columns["E"][rowActual].Value = int.Parse(row["CuentaSAPCodigo"].ToString());
                            worksheetMovimiento.Columns["F"][rowActual].Value = row["CuentaSAPN"].ToString();
                            worksheetMovimiento.Columns["G"][rowActual].Value = row["ArticuloGrupoCodigo"].ToString();
                            worksheetMovimiento.Columns["H"][rowActual].Value = row["ArticuloGrupoN"].ToString();
                            worksheetMovimiento.Columns["I"][rowActual].Value = row["MaterialCodigo"].ToString();
                            worksheetMovimiento.Columns["J"][rowActual].Value = row["MaterialN"].ToString();
                            worksheetMovimiento.Columns["K"][rowActual].Value = row["Anio_Total"].ToString();

                            double m1 = 0;
                            double.TryParse(row["M1"].ToString(), out m1);
                            double m2 = 0;
                            double.TryParse(row["M2"].ToString(), out m2);
                            double m3 = 0;
                            double.TryParse(row["M3"].ToString(), out m3);
                            double m4 = 0;
                            double.TryParse(row["M4"].ToString(), out m4);
                            double m5 = 0;
                            double.TryParse(row["M5"].ToString(), out m5);
                            double m6 = 0;
                            double.TryParse(row["M6"].ToString(), out m6);
                            double m7 = 0;
                            double.TryParse(row["M7"].ToString(), out m7);
                            double m8 = 0;
                            double.TryParse(row["M8"].ToString(), out m8);
                            double m9 = 0;
                            double.TryParse(row["M9"].ToString(), out m9);
                            double m10 = 0;
                            double.TryParse(row["M10"].ToString(), out m10);
                            double m11 = 0;
                            double.TryParse(row["M11"].ToString(), out m11);
                            double m12 = 0;
                            double.TryParse(row["M12"].ToString(), out m12);

                            double total = 0;
                            double.TryParse(row["Total"].ToString(), out total);

                            int valorK = 0;
                            if (worksheetMovimiento.Columns["K"][rowActual].Value != null)
                            {
                                string valorKStr = worksheetMovimiento.Columns["K"][rowActual].Value.ToString();
                                if (!string.IsNullOrEmpty(valorKStr) && valorKStr != "0" && valorKStr != "-")
                                {
                                    int.TryParse(valorKStr, out valorK);
                                }
                            }

                            if (valorK - periodo == 1)
                            {
                                worksheetMovimiento.Columns["L"][rowActual].Value = m1;
                                worksheetMovimiento.Columns["M"][rowActual].Value = m2;
                                worksheetMovimiento.Columns["N"][rowActual].Value = m3;
                                worksheetMovimiento.Columns["O"][rowActual].Value = m4;
                                worksheetMovimiento.Columns["P"][rowActual].Value = m5;
                                worksheetMovimiento.Columns["Q"][rowActual].Value = m6;
                                worksheetMovimiento.Columns["R"][rowActual].Value = m7;
                                worksheetMovimiento.Columns["S"][rowActual].Value = m8;
                                worksheetMovimiento.Columns["T"][rowActual].Value = m9;
                                worksheetMovimiento.Columns["U"][rowActual].Value = m10;
                                worksheetMovimiento.Columns["V"][rowActual].Value = m11;
                                worksheetMovimiento.Columns["W"][rowActual].Value = m12;
                                worksheetMovimiento.Columns["X"][rowActual].Value = total;

                                double prom  = total/12;

                                worksheetMovimiento.Columns["Y"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["Z"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AA"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AB"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AC"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AD"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AE"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AF"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AG"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AH"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AI"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AJ"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AK"][rowActual].Value = total;

                            }
                            else if (valorK - periodo == 2)
                            {
                                worksheetMovimiento.Columns["L"][rowActual].Value = m1;
                                worksheetMovimiento.Columns["M"][rowActual].Value = m2;
                                worksheetMovimiento.Columns["N"][rowActual].Value = m3;
                                worksheetMovimiento.Columns["O"][rowActual].Value = m4;
                                worksheetMovimiento.Columns["P"][rowActual].Value = m5;
                                worksheetMovimiento.Columns["Q"][rowActual].Value = m6;
                                worksheetMovimiento.Columns["R"][rowActual].Value = m7;
                                worksheetMovimiento.Columns["S"][rowActual].Value = m8;
                                worksheetMovimiento.Columns["T"][rowActual].Value = m9;
                                worksheetMovimiento.Columns["U"][rowActual].Value = m10;
                                worksheetMovimiento.Columns["V"][rowActual].Value = m11;
                                worksheetMovimiento.Columns["W"][rowActual].Value = m12;
                                worksheetMovimiento.Columns["X"][rowActual].Value = total;

                                double prom = total / 12;

                                worksheetMovimiento.Columns["Y"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["Z"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AA"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AB"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AC"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AD"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AE"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AF"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AG"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AH"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AI"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AJ"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AK"][rowActual].Value = total;

                                worksheetMovimiento.Columns["AL"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AM"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AN"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AO"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AP"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AQ"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AR"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AS"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AT"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AU"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AV"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AW"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AX"][rowActual].Value = total;
                            }
                            else
                            {
                                worksheetMovimiento.Columns["L"][rowActual].Value = m1;
                                worksheetMovimiento.Columns["M"][rowActual].Value = m2;
                                worksheetMovimiento.Columns["N"][rowActual].Value = m3;
                                worksheetMovimiento.Columns["O"][rowActual].Value = m4;
                                worksheetMovimiento.Columns["P"][rowActual].Value = m5;
                                worksheetMovimiento.Columns["Q"][rowActual].Value = m6;
                                worksheetMovimiento.Columns["R"][rowActual].Value = m7;
                                worksheetMovimiento.Columns["S"][rowActual].Value = m8;
                                worksheetMovimiento.Columns["T"][rowActual].Value = m9;
                                worksheetMovimiento.Columns["U"][rowActual].Value = m10;
                                worksheetMovimiento.Columns["V"][rowActual].Value = m11;
                                worksheetMovimiento.Columns["W"][rowActual].Value = m12;
                                worksheetMovimiento.Columns["X"][rowActual].Value = total;
                            }


                            rowActual++;
                        }

                        int lastRow = rowActual + 1;

                        string finalMergeRange = String.Format("L{0}:X{0}", lastRow);
                        worksheetMovimiento.Range[finalMergeRange].Merge();
                        worksheetMovimiento.Cells[String.Format("L{0}", lastRow)].Value = periodo.ToString();

                        int periodo1 = periodo + 1; 

                        string finalMergeRange2 = String.Format("Y{0}:AK{0}", lastRow);
                        worksheetMovimiento.Range[finalMergeRange2].Merge();
                        worksheetMovimiento.Cells[String.Format("Y{0}", lastRow)].Value = periodo1.ToString();

                        int periodo2 = periodo + 2;

                        string finalMergeRange3 = String.Format("AL{0}:AX{0}", lastRow);
                        worksheetMovimiento.Range[finalMergeRange3].Merge();
                        worksheetMovimiento.Cells[String.Format("AL{0}", lastRow)].Value = periodo2 .ToString();

                        worksheetMovimiento.Range[finalMergeRange].Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center;
                        worksheetMovimiento.Range[finalMergeRange].Alignment.Vertical = DevExpress.Spreadsheet.SpreadsheetVerticalAlignment.Center;

                        worksheetMovimiento.Range[finalMergeRange2].Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center;
                        worksheetMovimiento.Range[finalMergeRange2].Alignment.Vertical = DevExpress.Spreadsheet.SpreadsheetVerticalAlignment.Center;

                        worksheetMovimiento.Range[finalMergeRange3].Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center;
                        worksheetMovimiento.Range[finalMergeRange3].Alignment.Vertical = DevExpress.Spreadsheet.SpreadsheetVerticalAlignment.Center;

                        worksheetMovimiento.Cells[String.Format("L{0}", lastRow)].NumberFormat = "@";
                        worksheetMovimiento.Cells[String.Format("Y{0}", lastRow)].NumberFormat = "@";
                        worksheetMovimiento.Cells[String.Format("AL{0}", lastRow)].NumberFormat = "@";

                        worksheetMovimiento.Range[finalMergeRange].Font.Bold = true;
                        worksheetMovimiento.Range[finalMergeRange].Fill.BackgroundColor = rowColor;
                        worksheetMovimiento.Range[finalMergeRange].Font.Color = rowColorF;

                        worksheetMovimiento.Range[finalMergeRange2].Font.Bold = true;
                        worksheetMovimiento.Range[finalMergeRange2].Fill.BackgroundColor = rowColor;
                        worksheetMovimiento.Range[finalMergeRange2].Font.Color = rowColorF;

                        worksheetMovimiento.Range[finalMergeRange3].Font.Bold = true;
                        worksheetMovimiento.Range[finalMergeRange3].Fill.BackgroundColor = rowColor;
                        worksheetMovimiento.Range[finalMergeRange3].Font.Color = rowColorF;

                        worksheetMovimiento.Range["L:AX"].NumberFormat = "#,##0.00";

                        worksheetMovimiento.Range["A1:AX" + (rowActual)].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);
                    }
                    else if (maxAnioTotal - periodo == 1)
                    {
                        worksheetMovimiento.Cells["A1"].Value = "CEBE";
                        worksheetMovimiento.Cells["B1"].Value = "CEBE NOMBRE";
                        worksheetMovimiento.Cells["C1"].Value = "CECO";
                        worksheetMovimiento.Cells["D1"].Value = "CECO NOMBRE";
                        worksheetMovimiento.Cells["E1"].Value = "CUENTA";
                        worksheetMovimiento.Cells["F1"].Value = "CUENTA NOMBRE";
                        worksheetMovimiento.Cells["G1"].Value = "GRUPO DE ARTICULO";
                        worksheetMovimiento.Cells["H1"].Value = "GRUPO DE ARTICULO NOMBRE";
                        worksheetMovimiento.Cells["I1"].Value = "CODIGO DE MATERIAL";
                        worksheetMovimiento.Cells["J1"].Value = "CODIGO DE MATERIAL NOMBRE";
                        worksheetMovimiento.Cells["K1"].Value = "AÑO HASTA";

                        worksheetMovimiento.Cells["L1"].Value = "ENERO";
                        worksheetMovimiento.Cells["M1"].Value = "FEBRERO";
                        worksheetMovimiento.Cells["N1"].Value = "MARZO";
                        worksheetMovimiento.Cells["O1"].Value = "ABRIL";
                        worksheetMovimiento.Cells["P1"].Value = "MAYO";
                        worksheetMovimiento.Cells["Q1"].Value = "JUNIO";
                        worksheetMovimiento.Cells["R1"].Value = "JULIO";
                        worksheetMovimiento.Cells["S1"].Value = "AGOSTO";
                        worksheetMovimiento.Cells["T1"].Value = "SEPTIEMBRE";
                        worksheetMovimiento.Cells["U1"].Value = "OCTUBRE";
                        worksheetMovimiento.Cells["V1"].Value = "NOVIEMBRE";
                        worksheetMovimiento.Cells["W1"].Value = "DICIEMBRE";
                        worksheetMovimiento.Cells["X1"].Value = "TOTAL";

                        worksheetMovimiento.Cells["Y1"].Value = "ENERO";
                        worksheetMovimiento.Cells["Z1"].Value = "FEBRERO";
                        worksheetMovimiento.Cells["AA1"].Value = "MARZO";
                        worksheetMovimiento.Cells["AB1"].Value = "ABRIL";
                        worksheetMovimiento.Cells["AC1"].Value = "MAYO";
                        worksheetMovimiento.Cells["AD1"].Value = "JUNIO";
                        worksheetMovimiento.Cells["AE1"].Value = "JULIO";
                        worksheetMovimiento.Cells["AF1"].Value = "AGOSTO";
                        worksheetMovimiento.Cells["AG1"].Value = "SEPTIEMBRE";
                        worksheetMovimiento.Cells["AH1"].Value = "OCTUBRE";
                        worksheetMovimiento.Cells["AI1"].Value = "NOVIEMBRE";
                        worksheetMovimiento.Cells["AJ1"].Value = "DICIEMBRE";
                        worksheetMovimiento.Cells["AK1"].Value = "TOTAL";

                        worksheetMovimiento.Range["A1:K1"].ColumnWidth = 75;
                        worksheetMovimiento.Cells["B1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["D1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["F1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["H1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["J1"].ColumnWidth = 170;

                        worksheetMovimiento.Range["L1:AK1"].ColumnWidth = 70;
                        worksheetMovimiento.Range["A1:AK1"].Alignment.WrapText = true;
                        worksheetMovimiento.Range["A1:AK1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                        worksheetMovimiento.Range["A1:AK1"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                        worksheetMovimiento.Range["A1:J1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                        worksheetMovimiento.Range["A1:AK1"].Font.Bold = true;
                        worksheetMovimiento.Range["A1:AK1"].Fill.BackgroundColor = rowColor;
                        worksheetMovimiento.Range["A1:AK1"].Font.Color = rowColorF;

                        //Detalle                
                        int rowActual = 1;
                        foreach (DataRow row in oData.Rows)
                        {

                            worksheetMovimiento.Columns["A"][rowActual].Value = int.Parse(row["CEBECodigo"].ToString());
                            worksheetMovimiento.Columns["B"][rowActual].Value = row["CEBEN"].ToString();
                            worksheetMovimiento.Columns["C"][rowActual].Value = row["CECOCodigo"].ToString();
                            worksheetMovimiento.Columns["D"][rowActual].Value = row["CECON"].ToString();
                            worksheetMovimiento.Columns["E"][rowActual].Value = int.Parse(row["CuentaSAPCodigo"].ToString());
                            worksheetMovimiento.Columns["F"][rowActual].Value = row["CuentaSAPN"].ToString();
                            worksheetMovimiento.Columns["G"][rowActual].Value = row["ArticuloGrupoCodigo"].ToString();
                            worksheetMovimiento.Columns["H"][rowActual].Value = row["ArticuloGrupoN"].ToString();
                            worksheetMovimiento.Columns["I"][rowActual].Value = row["MaterialCodigo"].ToString();
                            worksheetMovimiento.Columns["J"][rowActual].Value = row["MaterialN"].ToString();
                            worksheetMovimiento.Columns["K"][rowActual].Value = row["Anio_Total"].ToString();

                            double m1 = 0;
                            double.TryParse(row["M1"].ToString(), out m1);
                            double m2 = 0;
                            double.TryParse(row["M2"].ToString(), out m2);
                            double m3 = 0;
                            double.TryParse(row["M3"].ToString(), out m3);
                            double m4 = 0;
                            double.TryParse(row["M4"].ToString(), out m4);
                            double m5 = 0;
                            double.TryParse(row["M5"].ToString(), out m5);
                            double m6 = 0;
                            double.TryParse(row["M6"].ToString(), out m6);
                            double m7 = 0;
                            double.TryParse(row["M7"].ToString(), out m7);
                            double m8 = 0;
                            double.TryParse(row["M8"].ToString(), out m8);
                            double m9 = 0;
                            double.TryParse(row["M9"].ToString(), out m9);
                            double m10 = 0;
                            double.TryParse(row["M10"].ToString(), out m10);
                            double m11 = 0;
                            double.TryParse(row["M11"].ToString(), out m11);
                            double m12 = 0;
                            double.TryParse(row["M12"].ToString(), out m12);

                            double total = 0;
                            double.TryParse(row["Total"].ToString(), out total);


                            int valorK = 0;
                            if (worksheetMovimiento.Columns["K"][rowActual].Value != null)
                            {
                                string valorKStr = worksheetMovimiento.Columns["K"][rowActual].Value.ToString();
                                if (!string.IsNullOrEmpty(valorKStr) && valorKStr != "0" && valorKStr != "-")
                                {
                                    int.TryParse(valorKStr, out valorK);
                                }
                            }

                            if (valorK - maxAnioTotal == 1)
                            {
                                worksheetMovimiento.Columns["L"][rowActual].Value = m1;
                                worksheetMovimiento.Columns["M"][rowActual].Value = m2;
                                worksheetMovimiento.Columns["N"][rowActual].Value = m3;
                                worksheetMovimiento.Columns["O"][rowActual].Value = m4;
                                worksheetMovimiento.Columns["P"][rowActual].Value = m5;
                                worksheetMovimiento.Columns["Q"][rowActual].Value = m6;
                                worksheetMovimiento.Columns["R"][rowActual].Value = m7;
                                worksheetMovimiento.Columns["S"][rowActual].Value = m8;
                                worksheetMovimiento.Columns["T"][rowActual].Value = m9;
                                worksheetMovimiento.Columns["U"][rowActual].Value = m10;
                                worksheetMovimiento.Columns["V"][rowActual].Value = m11;
                                worksheetMovimiento.Columns["W"][rowActual].Value = m12;
                                worksheetMovimiento.Columns["X"][rowActual].Value = total;

                                double prom = (m1 + m2 + m3 + m4 + m5 + m6 + m7 + m8 + m9 + m10 + m11 + m12) / total;

                                worksheetMovimiento.Columns["Y"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["Z"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AA"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AB"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AC"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AD"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AE"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AF"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AG"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AH"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AI"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AJ"][rowActual].Value = prom;
                                worksheetMovimiento.Columns["AK"][rowActual].Value = total;

                            }
                            else
                            {
                                worksheetMovimiento.Columns["L"][rowActual].Value = m1;
                                worksheetMovimiento.Columns["M"][rowActual].Value = m2;
                                worksheetMovimiento.Columns["N"][rowActual].Value = m3;
                                worksheetMovimiento.Columns["O"][rowActual].Value = m4;
                                worksheetMovimiento.Columns["P"][rowActual].Value = m5;
                                worksheetMovimiento.Columns["Q"][rowActual].Value = m6;
                                worksheetMovimiento.Columns["R"][rowActual].Value = m7;
                                worksheetMovimiento.Columns["S"][rowActual].Value = m8;
                                worksheetMovimiento.Columns["T"][rowActual].Value = m9;
                                worksheetMovimiento.Columns["U"][rowActual].Value = m10;
                                worksheetMovimiento.Columns["V"][rowActual].Value = m11;
                                worksheetMovimiento.Columns["W"][rowActual].Value = m12;
                                worksheetMovimiento.Columns["X"][rowActual].Value = total;
                            }


                            rowActual++;
                        }

                        int lastRow = rowActual + 1;

                        string finalMergeRange = String.Format("L{0}:X{0}", lastRow);
                        worksheetMovimiento.Range[finalMergeRange].Merge();
                        worksheetMovimiento.Cells[String.Format("L{0}", lastRow)].Value = periodo.ToString();

                        int periodo1 = periodo + 1;

                        string finalMergeRange2 = String.Format("Y{0}:AK{0}", lastRow);
                        worksheetMovimiento.Range[finalMergeRange2].Merge();
                        worksheetMovimiento.Cells[String.Format("Y{0}", lastRow)].Value = periodo1.ToString();

                        worksheetMovimiento.Range[finalMergeRange].Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center;
                        worksheetMovimiento.Range[finalMergeRange].Alignment.Vertical = DevExpress.Spreadsheet.SpreadsheetVerticalAlignment.Center;

                        worksheetMovimiento.Range[finalMergeRange2].Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center;
                        worksheetMovimiento.Range[finalMergeRange2].Alignment.Vertical = DevExpress.Spreadsheet.SpreadsheetVerticalAlignment.Center;

                        worksheetMovimiento.Cells[String.Format("L{0}", lastRow)].NumberFormat = "@";
                        worksheetMovimiento.Cells[String.Format("Y{0}", lastRow)].NumberFormat = "@";

                        worksheetMovimiento.Range[finalMergeRange].Font.Bold = true;
                        worksheetMovimiento.Range[finalMergeRange].Fill.BackgroundColor = rowColor;
                        worksheetMovimiento.Range[finalMergeRange].Font.Color = rowColorF;

                        worksheetMovimiento.Range[finalMergeRange2].Font.Bold = true;
                        worksheetMovimiento.Range[finalMergeRange2].Fill.BackgroundColor = rowColor;
                        worksheetMovimiento.Range[finalMergeRange2].Font.Color = rowColorF;

                        worksheetMovimiento.Range["L:AX"].NumberFormat = "#,##0.00";

                        worksheetMovimiento.Range["A1:AK" + (rowActual)].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);

                    }
                    else
                    {
                        //Cabecera
                        worksheetMovimiento.Cells["A1"].Value = "CEBE";
                        worksheetMovimiento.Cells["B1"].Value = "CEBE NOMBRE";
                        worksheetMovimiento.Cells["C1"].Value = "CECO";
                        worksheetMovimiento.Cells["D1"].Value = "CECO NOMBRE";
                        worksheetMovimiento.Cells["E1"].Value = "CUENTA";
                        worksheetMovimiento.Cells["F1"].Value = "CUENTA NOMBRE";
                        worksheetMovimiento.Cells["G1"].Value = "GRUPO DE ARTICULO";
                        worksheetMovimiento.Cells["H1"].Value = "GRUPO DE ARTICULO NOMBRE";
                        worksheetMovimiento.Cells["I1"].Value = "CODIGO DE MATERIAL";
                        worksheetMovimiento.Cells["J1"].Value = "CODIGO DE MATERIAL NOMBRE";
                        worksheetMovimiento.Cells["K1"].Value = "AÑO HASTA";

                        worksheetMovimiento.Cells["L1"].Value = "ENERO";
                        worksheetMovimiento.Cells["M1"].Value = "FEBRERO";
                        worksheetMovimiento.Cells["N1"].Value = "MARZO";
                        worksheetMovimiento.Cells["O1"].Value = "ABRIL";
                        worksheetMovimiento.Cells["P1"].Value = "MAYO";
                        worksheetMovimiento.Cells["Q1"].Value = "JUNIO";
                        worksheetMovimiento.Cells["R1"].Value = "JULIO";
                        worksheetMovimiento.Cells["S1"].Value = "AGOSTO";
                        worksheetMovimiento.Cells["T1"].Value = "SEPTIEMBRE";
                        worksheetMovimiento.Cells["U1"].Value = "OCTUBRE";
                        worksheetMovimiento.Cells["V1"].Value = "NOVIEMBRE";
                        worksheetMovimiento.Cells["W1"].Value = "DICIEMBRE";
                        worksheetMovimiento.Cells["X1"].Value = "TOTAL";

                        worksheetMovimiento.Range["A1:K1"].ColumnWidth = 75;
                        worksheetMovimiento.Cells["B1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["D1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["F1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["H1"].ColumnWidth = 170;
                        worksheetMovimiento.Cells["J1"].ColumnWidth = 170;

                        worksheetMovimiento.Range["L1:X1"].ColumnWidth = 70;
                        worksheetMovimiento.Range["A1:X1"].Alignment.WrapText = true;
                        worksheetMovimiento.Range["A1:X1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                        worksheetMovimiento.Range["A1:X1"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                        worksheetMovimiento.Range["A1:J1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                        worksheetMovimiento.Range["A1:X1"].Font.Bold = true;
                        worksheetMovimiento.Range["A1:X1"].Fill.BackgroundColor = rowColor;
                        worksheetMovimiento.Range["A1:X1"].Font.Color = rowColorF;

                        //Detalle                
                        int rowActual = 1;
                        foreach (DataRow row in oData.Rows)
                        {

                            worksheetMovimiento.Columns["A"][rowActual].Value = int.Parse(row["CEBECodigo"].ToString());
                            worksheetMovimiento.Columns["B"][rowActual].Value = row["CEBEN"].ToString();
                            worksheetMovimiento.Columns["C"][rowActual].Value = row["CECOCodigo"].ToString();
                            worksheetMovimiento.Columns["D"][rowActual].Value = row["CECON"].ToString();
                            worksheetMovimiento.Columns["E"][rowActual].Value = int.Parse(row["CuentaSAPCodigo"].ToString());
                            worksheetMovimiento.Columns["F"][rowActual].Value = row["CuentaSAPN"].ToString();
                            worksheetMovimiento.Columns["G"][rowActual].Value = row["ArticuloGrupoCodigo"].ToString();
                            worksheetMovimiento.Columns["H"][rowActual].Value = row["ArticuloGrupoN"].ToString();
                            worksheetMovimiento.Columns["I"][rowActual].Value = row["MaterialCodigo"].ToString();
                            worksheetMovimiento.Columns["J"][rowActual].Value = row["MaterialN"].ToString();
                            worksheetMovimiento.Columns["K"][rowActual].Value = row["Anio_Total"].ToString();

                            double m1 = 0;
                            double.TryParse(row["M1"].ToString(), out m1);
                            double m2 = 0;
                            double.TryParse(row["M2"].ToString(), out m2);
                            double m3 = 0;
                            double.TryParse(row["M3"].ToString(), out m3);
                            double m4 = 0;
                            double.TryParse(row["M4"].ToString(), out m4);
                            double m5 = 0;
                            double.TryParse(row["M5"].ToString(), out m5);
                            double m6 = 0;
                            double.TryParse(row["M6"].ToString(), out m6);
                            double m7 = 0;
                            double.TryParse(row["M7"].ToString(), out m7);
                            double m8 = 0;
                            double.TryParse(row["M8"].ToString(), out m8);
                            double m9 = 0;
                            double.TryParse(row["M9"].ToString(), out m9);
                            double m10 = 0;
                            double.TryParse(row["M10"].ToString(), out m10);
                            double m11 = 0;
                            double.TryParse(row["M11"].ToString(), out m11);
                            double m12 = 0;
                            double.TryParse(row["M12"].ToString(), out m12);

                            double total = 0;
                            double.TryParse(row["Total"].ToString(), out total);

                            worksheetMovimiento.Columns["L"][rowActual].Value = m1;
                            worksheetMovimiento.Columns["M"][rowActual].Value = m2;
                            worksheetMovimiento.Columns["N"][rowActual].Value = m3;
                            worksheetMovimiento.Columns["O"][rowActual].Value = m4;
                            worksheetMovimiento.Columns["P"][rowActual].Value = m5;
                            worksheetMovimiento.Columns["Q"][rowActual].Value = m6;
                            worksheetMovimiento.Columns["R"][rowActual].Value = m7;
                            worksheetMovimiento.Columns["S"][rowActual].Value = m8;
                            worksheetMovimiento.Columns["T"][rowActual].Value = m9;
                            worksheetMovimiento.Columns["U"][rowActual].Value = m10;
                            worksheetMovimiento.Columns["V"][rowActual].Value = m11;
                            worksheetMovimiento.Columns["W"][rowActual].Value = m12;
                            worksheetMovimiento.Columns["X"][rowActual].Value = total;

                            rowActual++;
                        }

                        int lastRow = rowActual + 1;

                        string finalMergeRange = String.Format("L{0}:X{0}", lastRow);
                        worksheetMovimiento.Range[finalMergeRange].Merge();
                        worksheetMovimiento.Cells[String.Format("L{0}", lastRow)].Value = periodo.ToString();

                        worksheetMovimiento.Range[finalMergeRange].Alignment.Horizontal = DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center;
                        worksheetMovimiento.Range[finalMergeRange].Alignment.Vertical = DevExpress.Spreadsheet.SpreadsheetVerticalAlignment.Center;

                        worksheetMovimiento.Cells[String.Format("L{0}", lastRow)].NumberFormat = "@";

                        worksheetMovimiento.Range[finalMergeRange].Font.Bold = true;
                        worksheetMovimiento.Range[finalMergeRange].Fill.BackgroundColor = rowColor;
                        worksheetMovimiento.Range[finalMergeRange].Font.Color = rowColorF;

                        worksheetMovimiento.Range["L:X"].NumberFormat = "#,##0.00";

                        worksheetMovimiento.Range["A1:X" + (rowActual)].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);
                    }

                }
                finally
                {
                    workbook.EndUpdate();
                }

                string nombreExcel = VersionDDL.Text + " - " + PeriodoTxt.Text + " (Horizontal).xlsx";
                //// Save the document file under the specified name.
                workbook.SaveDocument(Server.MapPath(nombreExcel), DocumentFormat.OpenXml);

                FileInfo myfile = new FileInfo(Server.MapPath(nombreExcel));
                // Clear the content of the response
                byte[] archivo;
                Response.ClearContent();
                Response.Clear();
                // Set the ContentType
                Response.Cache.SetCacheability(HttpCacheability.Private);
                Response.ContentType = "application/vnd.ms-excel";//application/vnd.ms-excel   //application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
                                                                  // Add the file name and attachment, which will force the open/cancel/save dialog box to show, to the header
                Response.AddHeader("Content-Disposition", "attachment; filename=" + nombreExcel);
                // Add the file size into the response header
                Response.AddHeader("Content-Length", (myfile.Length).ToString());
                // Write the file into the response (TransmitFile is for ASP.NET 2.0. In ASP.NET 1.1 you have to use WriteFile instead)
                archivo = File.ReadAllBytes(myfile.FullName);
                Response.BinaryWrite(archivo);

                File.Delete(Server.MapPath(nombreExcel));

                Response.End();

            }
        }
    }

    public void ExportarInfoVertical()
    {
        Base oBase = new Base();

        int version = int.Parse(VersionDDL.Value.ToString());

        Color rowColor = System.Drawing.ColorTranslator.FromHtml("#167b67");
        Color rowColorF = System.Drawing.ColorTranslator.FromHtml("#ffffff");
        DateTime fechaMin = DateTime.Parse("01-01-1900");
        // Create a new workbook.
        using (Workbook workbook = new Workbook())
        {

            // Access the first worksheet in the workbook.


            // Set the unit of measurement.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;

            #region Movimientos
            DataTable oData = new PresupuestoVersionDataDAL(oBase.StrConexionPresupuestoAlerta).ExportarByFK(int.Parse(HttpContext.Current.User.Identity.Name), version, 0);
            Worksheet worksheetMovimiento = workbook.Worksheets[0];
            worksheetMovimiento.Name = "Plantilla Carga";
            workbook.BeginUpdate();
            try
            {
                worksheetMovimiento.OutlineOptions.SummaryRowsBelow = false;

                //Cabecera
                worksheetMovimiento.Cells["A1"].Value = "CEBE";
                worksheetMovimiento.Cells["B1"].Value = "CEBE NOMBRE";
                worksheetMovimiento.Cells["C1"].Value = "CECO";
                worksheetMovimiento.Cells["D1"].Value = "CECO NOMBRE";
                worksheetMovimiento.Cells["E1"].Value = "CUENTA";
                worksheetMovimiento.Cells["F1"].Value = "CUENTA NOMBRE";
                worksheetMovimiento.Cells["G1"].Value = "GRUPO DE ARTICULO";
                worksheetMovimiento.Cells["H1"].Value = "GRUPO DE ARTICULO NOMBRE";
                worksheetMovimiento.Cells["I1"].Value = "CODIGO DE MATERIAL";
                worksheetMovimiento.Cells["J1"].Value = "CODIGO DE MATERIAL NOMBRE";

                worksheetMovimiento.Cells["K1"].Value = "PERIODO";
                worksheetMovimiento.Cells["L1"].Value = "PERIODO NOMBRE";
                worksheetMovimiento.Cells["M1"].Value = "MONTO";
                worksheetMovimiento.Cells["N1"].Value = "AÑO HASTA";

                worksheetMovimiento.Range["A1:J1"].ColumnWidth = 75;
                worksheetMovimiento.Cells["B1"].ColumnWidth = 170;
                worksheetMovimiento.Cells["D1"].ColumnWidth = 170;
                worksheetMovimiento.Cells["F1"].ColumnWidth = 170;
                worksheetMovimiento.Cells["H1"].ColumnWidth = 170;
                worksheetMovimiento.Cells["J1"].ColumnWidth = 170;

                worksheetMovimiento.Range["K1:N1"].ColumnWidth = 80;
                worksheetMovimiento.Range["A1:N1"].Alignment.WrapText = true;
                worksheetMovimiento.Range["A1:N1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheetMovimiento.Range["A1:N1"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheetMovimiento.Range["A1:N1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                worksheetMovimiento.Range["A1:N1"].Font.Bold = true;
                worksheetMovimiento.Range["A1:N1"].Fill.BackgroundColor = rowColor;
                worksheetMovimiento.Range["A1:N1"].Font.Color = rowColorF;

                //Detalle                
                int rowActual = 1;
                for (int mes = 1; mes < 13; mes++)
                {
                    string nombreMes = "";
                    if (mes == 1) nombreMes = "ENERO";
                    if (mes == 2) nombreMes = "FEBRERO";
                    if (mes == 3) nombreMes = "MARZO";
                    if (mes == 4) nombreMes = "ABRIL";
                    if (mes == 5) nombreMes = "MAYO";
                    if (mes == 6) nombreMes = "JUNIO";
                    if (mes == 7) nombreMes = "JULIO";
                    if (mes == 8) nombreMes = "AGOSTO";
                    if (mes == 9) nombreMes = "SEPTIEMBRE";
                    if (mes == 10) nombreMes = "OCTUBRE";
                    if (mes == 11) nombreMes = "NOVIEMBRE";
                    if (mes == 12) nombreMes = "DICIEMBRE";


                    foreach (DataRow row in oData.Rows)
                    {

                        worksheetMovimiento.Columns["A"][rowActual].Value = int.Parse(row["CEBECodigo"].ToString());
                        worksheetMovimiento.Columns["B"][rowActual].Value = row["CEBEN"].ToString();
                        worksheetMovimiento.Columns["C"][rowActual].Value = row["CECOCodigo"].ToString();
                        worksheetMovimiento.Columns["D"][rowActual].Value = row["CECON"].ToString();
                        worksheetMovimiento.Columns["E"][rowActual].Value = int.Parse(row["CuentaSAPCodigo"].ToString());
                        worksheetMovimiento.Columns["F"][rowActual].Value = row["CuentaSAPN"].ToString();
                        worksheetMovimiento.Columns["G"][rowActual].Value = row["ArticuloGrupoCodigo"].ToString();
                        worksheetMovimiento.Columns["H"][rowActual].Value = row["ArticuloGrupoN"].ToString();
                        worksheetMovimiento.Columns["I"][rowActual].Value = row["MaterialCodigo"].ToString();
                        worksheetMovimiento.Columns["J"][rowActual].Value = row["MaterialN"].ToString();

                        double mesValor = 0;
                        double.TryParse(row["M"+ mes].ToString(), out mesValor);

                        worksheetMovimiento.Columns["K"][rowActual].Value = int.Parse(PeriodoTxt.Text) * 100 + mes;
                        worksheetMovimiento.Columns["L"][rowActual].Value = nombreMes + " " + PeriodoTxt.Text;
                        worksheetMovimiento.Columns["M"][rowActual].Value = mesValor;
                        worksheetMovimiento.Columns["N"][rowActual].Value = row["Anio_Total"].ToString();

                        rowActual++;
                    }
                }


                worksheetMovimiento.Range["M:M"].NumberFormat = "#,##0.00";

                worksheetMovimiento.Range["A1:N" + (rowActual)].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);

                // Enable filtering for the specified cell range.
                CellRange range = worksheetMovimiento.Range["A1:N" + (rowActual)];
                worksheetMovimiento.AutoFilter.Apply(range);

                //for (int i = 0; i < lines.Count; i++)
                //{
                //    int fin = lines.Count > (i + 1) ? lines[i + 1] : rowActual + 1;
                //    worksheetMovimiento.Rows.Group(lines[i], (fin - 2), false);
                //}

                //for (int i = 0; i < linesSub.Count; i++)
                //{
                //    int fin = linesSub.Count > (i + 1) ? linesSub[i + 1] : rowActual + 1;
                //    worksheetMovimiento.Rows.Group(linesSub[i], (fin - 2), false);
                //}

            }
            finally
            {
                workbook.EndUpdate();
            }

            #endregion




            // Calculate the workbook.
            //workbook.Calculate();








            //workbook


            string nombreExcel = VersionDDL.Text + " - " + PeriodoTxt.Text + " (Vertical).xlsx";
            //// Save the document file under the specified name.
            workbook.SaveDocument(Server.MapPath(nombreExcel), DocumentFormat.OpenXml);

            FileInfo myfile = new FileInfo(Server.MapPath(nombreExcel));
            // Clear the content of the response
            byte[] archivo;
            Response.ClearContent();
            Response.Clear();
            // Set the ContentType
            Response.Cache.SetCacheability(HttpCacheability.Private);
            Response.ContentType = "application/vnd.ms-excel";//application/vnd.ms-excel   //application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
                                                              // Add the file name and attachment, which will force the open/cancel/save dialog box to show, to the header
            Response.AddHeader("Content-Disposition", "attachment; filename=" + nombreExcel);
            // Add the file size into the response header
            Response.AddHeader("Content-Length", (myfile.Length).ToString());
            // Write the file into the response (TransmitFile is for ASP.NET 2.0. In ASP.NET 1.1 you have to use WriteFile instead)
            archivo = File.ReadAllBytes(myfile.FullName);
            Response.BinaryWrite(archivo);

            //Response.Flush();
            File.Delete(Server.MapPath(nombreExcel));


            //Response.WriteFile(myfile.FullName);
            // End the response
            Response.End();


            //// Export the document to PDF.
            //workbook.ExportToPdf("TestDoc.pdf");
        }
        // Open the PDF document using the default viewer.
        //System.Diagnostics.Process.Start("TestDoc.pdf");

        // Open the XLSX document using the default application.
        //System.Diagnostics.Process.Start("TestDoc.xlsx");
    }

    public void ExportarPlantillaCarga()
    {
        Base oBase = new Base();

        int version = int.Parse(VersionDDL.Value.ToString());

        //Color rowColorEntrada = System.Drawing.ColorTranslator.FromHtml("#e8ffe8");
        //Color rowColorSalida = System.Drawing.ColorTranslator.FromHtml("#ffe1e4");

        Color rowColor = System.Drawing.ColorTranslator.FromHtml("#167b67");
        Color rowColorF = System.Drawing.ColorTranslator.FromHtml("#ffffff");
        //Color movimientoColor = System.Drawing.ColorTranslator.FromHtml("#edf5ff");
        //Color saldoInicialColor = System.Drawing.ColorTranslator.FromHtml("#c3d4e6");
        //Color saldoDisponibleColor = System.Drawing.ColorTranslator.FromHtml("#e4e4e4");
        //Color saldoFinalColor = System.Drawing.ColorTranslator.FromHtml("#97afca");
        DateTime fechaMin = DateTime.Parse("01-01-1900");
        // Create a new workbook.
        using (Workbook workbook = new Workbook())
        {

            // Access the first worksheet in the workbook.


            // Set the unit of measurement.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;

            #region Movimientos
            DataTable oData = new PresupuestoVersionDataDAL(oBase.StrConexionPresupuestoAlerta).ExportarByFK(int.Parse(HttpContext.Current.User.Identity.Name), version, 1);
            Worksheet worksheetMovimiento = workbook.Worksheets[0];
            worksheetMovimiento.Name = "Plantilla Carga";
            workbook.BeginUpdate();
            try
            {
                worksheetMovimiento.OutlineOptions.SummaryRowsBelow = false;

                //Cabecera
                worksheetMovimiento.Cells["A1"].Value = "CEBE";
                worksheetMovimiento.Cells["B1"].Value = "CEBE NOMBRE";
                worksheetMovimiento.Cells["C1"].Value = "CECO";
                worksheetMovimiento.Cells["D1"].Value = "CECO NOMBRE";
                worksheetMovimiento.Cells["E1"].Value = "CUENTA";
                worksheetMovimiento.Cells["F1"].Value = "CUENTA NOMBRE";
                worksheetMovimiento.Cells["G1"].Value = "GRUPO DE ARTICULO";
                worksheetMovimiento.Cells["H1"].Value = "GRUPO DE ARTICULO NOMBRE";
                worksheetMovimiento.Cells["I1"].Value = "CODIGO DE MATERIAL";
                worksheetMovimiento.Cells["J1"].Value = "CODIGO DE MATERIAL NOMBRE";
                //DESARROLLO DUHOVIT GRC ADEMAS SE REORDENARON LAS FILAS PARA MOSTRAR LA INFORMACION EN LA EXPORTACION 
                worksheetMovimiento.Cells["K1"].Value = "AÑO HASTA";
                //DESARROLLO DUHOVIT GRC
                worksheetMovimiento.Cells["L1"].Value = "ENERO";
                worksheetMovimiento.Cells["M1"].Value = "FEBRERO";
                worksheetMovimiento.Cells["N1"].Value = "MARZO";
                worksheetMovimiento.Cells["O1"].Value = "ABRIL";
                worksheetMovimiento.Cells["P1"].Value = "MAYO";
                worksheetMovimiento.Cells["Q1"].Value = "JUNIO";
                worksheetMovimiento.Cells["R1"].Value = "JULIO";
                worksheetMovimiento.Cells["S1"].Value = "AGOSTO";
                worksheetMovimiento.Cells["T1"].Value = "SEPTIEMBRE";
                worksheetMovimiento.Cells["U1"].Value = "OCTUBRE";
                worksheetMovimiento.Cells["V1"].Value = "NOVIEMBRE";
                worksheetMovimiento.Cells["W1"].Value = "DICIEMBRE";

                worksheetMovimiento.Range["A1:K1"].ColumnWidth = 100;
                worksheetMovimiento.Range["L1:W1"].ColumnWidth = 70;
                worksheetMovimiento.Range["A1:W1"].Alignment.WrapText = true;
                worksheetMovimiento.Range["A1:W1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheetMovimiento.Range["A1:W1"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheetMovimiento.Range["A1:W1"].Font.Bold = true;
                worksheetMovimiento.Range["A1:W1"].Fill.BackgroundColor = rowColor;
                worksheetMovimiento.Range["A1:W1"].Font.Color = rowColorF;

                //Detalle                
                int rowActual = 1;
                int colActual = 0;
                foreach (DataRow row in oData.Rows)
                {

                    worksheetMovimiento.Columns["A"][rowActual].Value = int.Parse(row["CEBECodigo"].ToString());
                    worksheetMovimiento.Columns["B"][rowActual].Value = row["CEBEN"].ToString();
                    worksheetMovimiento.Columns["C"][rowActual].Value = row["CECOCodigo"].ToString();
                    worksheetMovimiento.Columns["D"][rowActual].Value = row["CECON"].ToString();
                    worksheetMovimiento.Columns["E"][rowActual].Value = int.Parse(row["CuentaSAPCodigo"].ToString());
                    worksheetMovimiento.Columns["F"][rowActual].Value = row["CuentaSAPN"].ToString();
                    worksheetMovimiento.Columns["G"][rowActual].Value = row["ArticuloGrupoCodigo"].ToString();
                    worksheetMovimiento.Columns["H"][rowActual].Value = row["ArticuloGrupoN"].ToString();
                    worksheetMovimiento.Columns["I"][rowActual].Value = row["MaterialCodigo"].ToString();
                    worksheetMovimiento.Columns["J"][rowActual].Value = row["MaterialN"].ToString();
                    worksheetMovimiento.Columns["K"][rowActual].Value = 0;//Año hasta

                    worksheetMovimiento.Columns["L"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["M"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["N"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["O"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["P"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["Q"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["R"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["S"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["T"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["U"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["V"][rowActual].Value = 0;
                    worksheetMovimiento.Columns["W"][rowActual].Value = 0;

                    rowActual++;
                }


                worksheetMovimiento.Range["L:W"].NumberFormat = "#,##0.00";
                worksheetMovimiento.Range["A1:W" + (rowActual)].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);

                //// Enable filtering for the specified cell range.
                //CellRange range = worksheetMovimiento.Range["A1:O" + (rowActual)];
                //worksheetMovimiento.AutoFilter.Apply(range);

                //for (int i = 0; i < lines.Count; i++)
                //{
                //    int fin = lines.Count > (i + 1) ? lines[i + 1] : rowActual + 1;
                //    worksheetMovimiento.Rows.Group(lines[i], (fin - 2), false);
                //}

                //for (int i = 0; i < linesSub.Count; i++)
                //{
                //    int fin = linesSub.Count > (i + 1) ? linesSub[i + 1] : rowActual + 1;
                //    worksheetMovimiento.Rows.Group(linesSub[i], (fin - 2), false);
                //}

            }
            finally
            {
                workbook.EndUpdate();
            }

            #endregion




            // Calculate the workbook.
            //workbook.Calculate();








            //workbook


            string nombreExcel = "Planilla Base - "+ PeriodoTxt.Text + ".xlsx";
            //// Save the document file under the specified name.
            workbook.SaveDocument(Server.MapPath(nombreExcel), DocumentFormat.OpenXml);

            FileInfo myfile = new FileInfo(Server.MapPath(nombreExcel));
            // Clear the content of the response
            byte[] archivo;
            Response.ClearContent();
            Response.Clear();
            // Set the ContentType
            Response.Cache.SetCacheability(HttpCacheability.Private);
            Response.ContentType = "application/vnd.ms-excel";//application/vnd.ms-excel   //application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
                                                              // Add the file name and attachment, which will force the open/cancel/save dialog box to show, to the header
            Response.AddHeader("Content-Disposition", "attachment; filename=" + nombreExcel);
            // Add the file size into the response header
            Response.AddHeader("Content-Length", (myfile.Length).ToString());
            // Write the file into the response (TransmitFile is for ASP.NET 2.0. In ASP.NET 1.1 you have to use WriteFile instead)
            archivo = File.ReadAllBytes(myfile.FullName);
            Response.BinaryWrite(archivo);

            //Response.Flush();
            File.Delete(Server.MapPath(nombreExcel));


            //Response.WriteFile(myfile.FullName);
            // End the response
            Response.End();


            //// Export the document to PDF.
            //workbook.ExportToPdf("TestDoc.pdf");
        }
        // Open the PDF document using the default viewer.
        //System.Diagnostics.Process.Start("TestDoc.pdf");

        // Open the XLSX document using the default application.
        //System.Diagnostics.Process.Start("TestDoc.xlsx");
    }
}