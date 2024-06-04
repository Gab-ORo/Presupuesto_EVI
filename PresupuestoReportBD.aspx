<%@ Page Title="" Language="C#" MasterPageFile="~/SitePresupuestoAlerta.master" AutoEventWireup="true" CodeFile="PresupuestoReportBD.aspx.cs" Inherits="PresupuestoReportBD" %>

<%@ Register Assembly="DevExpress.Web.ASPxPivotGrid.v19.2, Version=19.2.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPivotGrid" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxTreeList.v19.2, Version=19.2.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxTreeList" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v19.2, Version=19.2.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web" TagPrefix="dx" %>


<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <style type="text/css">
        .dxgvDetailRow_Metropolis td.dxgvDetailCell_Metropolis {
            padding-top: 0;
        }

        .categoryTable {
            width: 100%;
        }

            .categoryTable .imageCell {
                padding: 2px;
            }

            .categoryTable .textCell {
                padding-left: 20px;
                width: 100%;
            }

        .textCell .label {
            color: #969696;
            text-transform: uppercase;
        }

        .textCell .description {
            font-size: 13px;
            width: 100%;
        }

        .captionCell {
            min-width: 120px !important;
            font-weight: bold;
        }

        .titulocss {
            padding: 0px;
            background-color: initial;
        }

        .progress-bar {
            background-color: #01a784;
        }

        .progress {
            background-color: #757575;
            font-weight: bold;
        }

        .dxgvDataRowHover_Metropolis a {
            color: #001aab !important;
        }

        [class*=dxtlNode_Metropolis]:hover {
            background-color: yellow;
        }
    </style>
    <div class="row" style="padding-top: 15px;">
        <div class="col-3">
            <dx:ASPxComboBox ID="CebeDDL" ClientInstanceName="cebeddl" runat="server"
                ValueType="System.String" Theme="Metropolis" Width="100%" Height="28px">
                <ClientSideEvents SelectedIndexChanged="function(s,e){Actualizar();}" />
            </dx:ASPxComboBox>
        </div>
        <div class="col-3">
            <dx:ASPxComboBox ID="CecoDDL" ClientInstanceName="cecoddl" runat="server"
                ValueType="System.String" Theme="Metropolis" Width="100%" Height="28px">
                <ClientSideEvents SelectedIndexChanged="function(s,e){Actualizar();}" />
            </dx:ASPxComboBox>
        </div>
        <div class="col-6">
            <div style="display: inline-flex;">
                <div>
                    <dx:ASPxDateEdit ID="PeriodoTxt" PickerType="Years" AllowUserInput="false"
                        ClientInstanceName="periodotxt" Width="80px" runat="server" Theme="Metropolis" Height="28px">
                        <ClientSideEvents DateChanged="function(s,e){ActualizarPeriodo();}" />
                        <CalendarProperties ShowTodayButton="false" ShowClearButton="false">
                        </CalendarProperties>
                    </dx:ASPxDateEdit>
                </div>
                <div style="padding-left: 5px;">
                    <dx:ASPxComboBox ID="VersionDDL" ClientInstanceName="versionddl" runat="server"
                        ValueType="System.String" Theme="Metropolis" Width="250" Height="28px" OnCallback="VersionDDL_Callback">
                        <ClientSideEvents SelectedIndexChanged="function(s,e){Actualizar();}" />
                    </dx:ASPxComboBox>
                </div>
            </div>
        </div>
    </div>
    <div class="row" style="padding-top: 15px;">
        <div class="col-12">
            <dx:ASPxTreeList ID="PresupuestoTreeList" ClientInstanceName="presupuestotreelist" runat="server" Theme="Metropolis"
                OnToolbarItemClick="PresupuestoTreeList_ToolbarItemClick"
                OnVirtualModeCreateChildren="PresupuestoTreeList_VirtualModeCreateChildren"
                OnVirtualModeNodeCreating="PresupuestoTreeList_VirtualModeNodeCreating"
                OnCustomCallback="PresupuestoTreeList_CustomCallback" Width="100%"
                AutoGenerateColumns="False">
                <Toolbars>
                    <dx:TreeListToolbar Enabled="true" Position="Top" ItemAlign="Left">
                        <Items>


                            <dx:TreeListToolbarItem Image-IconID="spreadsheet_expandfieldpivottable_16x16" Text="Agregar CEBE" Name="AddCebe"
                                ItemStyle-HoverStyle-BackColor="#FF8800" Visible="true" />

                            <dx:TreeListToolbarItem BeginGroup="true" Image-IconID="export_exporttoxlsx_16x16office2013" Text="Carga Presupuesto - Excel" Name="AddExcel"
                                ItemStyle-HoverStyle-BackColor="#FF8800" Visible="true" />

                            <dx:TreeListToolbarItem BeginGroup="true" Text="Eliminar Presupuesto" Name="DeletePre" Image-IconID="actions_cancel_16x16office2013"
                                ItemStyle-HoverStyle-BackColor="#FF8800" Visible="true" />

                            <dx:TreeListToolbarItem Command="Refresh" Image-IconID="actions_refresh_16x16office2013" Text="Actualizar" ItemStyle-HoverStyle-BackColor="#FF8800" />

                            <dx:TreeListToolbarItem Text="Exportar Información" Image-IconID="actions_download_16x16office2013" BeginGroup="true">
                                <Items>
                                    <dx:TreeListToolbarItem Name="ExportToXLSXH" Text="Horizontal" Image-IconID="export_exporttoxlsx_16x16office2013" />
                                    <dx:TreeListToolbarItem Name="ExportToXLSXV" Text="Vertical" Image-IconID="export_exporttoxlsx_16x16office2013" />
                                </Items>
                            </dx:TreeListToolbarItem>
                            <%-- BTN GRLL PLANTILLA DUHOVIT GABRIEL RODRIGUEZ --%>
                            <dx:TreeListToolbarItem Text="Plantilla de Carga" Image-IconID="actions_download_16x16office2013" BeginGroup="true">
                                <Items>
                                    <dx:TreeListToolbarItem Name="ExportToXLSXPC" Text="Plantilla" Image-IconID="export_exporttoxlsx_16x16office2013" />
                                </Items>
                            </dx:TreeListToolbarItem>

                            <dx:TreeListToolbarItem BeginGroup="true" Text="Ver Cargas Realizadas" Name="CargaPre" Image-IconID="actions_search_16x16devav"
                                ItemStyle-HoverStyle-BackColor="#FF8800" Visible="true" />

                        </Items>
                    </dx:TreeListToolbar>

                </Toolbars>
                <Columns>
                    <dx:TreeListDataColumn FieldName="Codigo" Caption="Código" CellStyle-Wrap="True" Width="80px" SortOrder="Ascending" SortIndex="0" />
                    <dx:TreeListDataColumn FieldName="EsUltimoNivel" Caption="+" CellStyle-Wrap="True" Width="25px" HeaderStyle-HorizontalAlign="Center">
                        <DataCellTemplate>
                            <i class="fas fa-plus-circle"
                                <%# 
                                    Eval("TieneSubNivel").ToString() == "1" && Eval("PuedePresupuestar").ToString() == "1" 
                                    ?
                                    "style='color: #f8ac59;cursor:pointer;' onclick='AgregarItem(" + Eval("Origen").ToString() + ",\""+Eval("CodigoIni").ToString()+"\");'"
                                    :
                                    "style='display:none;'"
                                %>></i>
                        </DataCellTemplate>
                    </dx:TreeListDataColumn>
                    <dx:TreeListDataColumn FieldName="EsUltimoNivel" Caption="-" CellStyle-Wrap="True" Width="25px" HeaderStyle-HorizontalAlign="Center">
                        <DataCellTemplate>

                            <%# Eval("PuedeEliminar").ToString() == "1" 
                                    ?
                                    "<i class='fas fa-trash-alt' style='color: red;cursor:pointer;' onclick='EliminarItem(\""+Eval("CodigoIni").ToString()+"\",\""+ Eval("Codigo").ToString()+" - "+Eval("Nombre").ToString() +"\","+Eval("Origen").ToString()+");' ></i>" : ""  %>
                        </DataCellTemplate>
                    </dx:TreeListDataColumn>
                    <dx:TreeListDataColumn FieldName="Nombre" Caption="Nombre" CellStyle-Wrap="True" FooterCellStyle-HorizontalAlign="Right" Width="360" />

                    <dx:TreeListDataColumn FieldName="" Name="" Caption="#" CellStyle-Wrap="True" Width="45px" HeaderStyle-HorizontalAlign="Center">
                        <DataCellTemplate>
                            <%# 
                                Eval("EsUltimoNivel").ToString() == "1" && Eval("PuedePresupuestar").ToString() == "1" 
                                ? 
                                (Eval("CuentaMadreCodigo").ToString() != "10000" &&  Eval("CuentaMadreCodigo").ToString() != "50000") ||  Eval("esAdmin").ToString() == "1" ?
                                      "<a style='color:blue;cursor:pointer;' onclick=\"EditarPre("+Eval("Origen")+",'"+ Eval("CodigoIni") +"');\">Editar</a>" 
                                    : ""
                                : ""  
                            %>
                        </DataCellTemplate>
                    </dx:TreeListDataColumn>

                    <%-- Desarrollo duhovit modificacion grilla GABRIEL RODRIGUEZ--%>

                    <dx:TreeListSpinEditColumn FieldName="Anio_Total" Name="Año Hasta" Caption="Año Hasta"
                        CellStyle-Wrap="True" EditCellStyle-HorizontalAlign="Center"
                        Width="100px" HeaderStyle-HorizontalAlign="Center">
                        <CellStyle VerticalAlign="Middle" HorizontalAlign="Center" />
                    </dx:TreeListSpinEditColumn>

                    <%-- Desarrollo duhovit modificacion grilla --%>
                    <dx:TreeListSpinEditColumn FieldName="Total" Name="Total" Caption="Total" CellStyle-Wrap="True"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        Width="100px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M1" Name="Enero" Caption="Enero"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M2" Name="Febrero" Caption="Febrero"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M3" Name="Marzo" Caption="Marzo"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M4" Name="Abril" Caption="Abril"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M5" Name="Mayo" Caption="Mayo"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M6" Name="Junio" Caption="Junio"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M7" Name="Julio" Caption="Julio"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M8" Name="Agosto" Caption="Agosto"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M9" Name="Septiembre" Caption="Septiembre"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M10" Name="Octubre" Caption="Octubre" PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M11" Name="Noviembre" Caption="Noviembre"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="M12" Name="Diciembre" Caption="Diciembre"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>
                    <dx:TreeListSpinEditColumn FieldName="Adicional" Name="Adicional" Caption="Adicional"
                        PropertiesSpinEdit-DisplayFormatString="{0:###,##0.#0}"
                        CellStyle-Wrap="True" Width="90px" HeaderStyle-HorizontalAlign="Center">
                    </dx:TreeListSpinEditColumn>

                </Columns>
                <Summary>
                    <dx:TreeListSummaryItem FieldName="Nombre" ShowInColumn="Nombre" SummaryType="Sum" DisplayFormat="Total" />
                    <dx:TreeListSummaryItem FieldName="Total" ShowInColumn="Total" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M1" ShowInColumn="M1" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M2" ShowInColumn="M2" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M3" ShowInColumn="M3" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M4" ShowInColumn="M4" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M5" ShowInColumn="M5" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M6" ShowInColumn="M6" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M7" ShowInColumn="M7" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M8" ShowInColumn="M8" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M9" ShowInColumn="M9" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M10" ShowInColumn="M10" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M11" ShowInColumn="M11" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />
                    <dx:TreeListSummaryItem FieldName="M12" ShowInColumn="M12" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />

                    <dx:TreeListSummaryItem FieldName="Adicional" ShowInColumn="Adicional" SummaryType="Sum" DisplayFormat="{0:###,##0.#0}" Recursive="false" />

                </Summary>
                <SettingsPager Mode="ShowAllNodes" EnableAdaptivity="true" />
                <Styles>
                    <AlternatingNode Enabled="true" />
                    <Footer BackColor="#005242" ForeColor="White" Font-Bold="true" />
                    <GroupFooter BackColor="#5d9489" ForeColor="White" Font-Bold="true" />
                    <FocusedNode BackColor="#788f8a"></FocusedNode>
                    <Node BackColor="White"></Node>
                </Styles>
                <Settings ShowGroupFooter="false" ShowFooter="True" GridLines="Both" ScrollableHeight="320" VerticalScrollBarMode="Visible" HorizontalScrollBarMode="Auto" />
                <SettingsBehavior ExpandCollapseAction="NodeDblClick" AllowFocusedNode="True" />
                <ClientSideEvents Init="function(s,e) {OnInit(s,e);}" EndCallback="function(s,e) {OnEndCallback(s,e);}"
                    ToolbarItemClick="function(s, e) {OnToolbarItemClick(s, e);}" />
            </dx:ASPxTreeList>
        </div>
    </div>

    <dx:ASPxGlobalEvents ID="ASPxGlobalEvents1" runat="server">
        <ClientSideEvents ControlsInitialized="OnControlsInitialized" />
    </dx:ASPxGlobalEvents>

    <asp:HiddenField ID="ReparticionAddHid" runat="server" />
    <asp:HiddenField ID="GerenciaAddHid" runat="server" />
    <asp:HiddenField ID="CentroCostoAddHid" runat="server" />
    <asp:HiddenField ID="CuentaContableAddHid" runat="server" />
    <asp:HiddenField ID="ImpactaAddHid" runat="server" />
    <asp:HiddenField ID="CuentaContableMadreAddHid" runat="server" />
    <asp:HiddenField ID="CuentaModuloAddHid" runat="server" />
    <asp:HiddenField ID="FuncionAddHid" runat="server" />
    <asp:HiddenField ID="FuncionCierreAddHid" runat="server" />

    <asp:HiddenField ID="VersionInfoHid" runat="server" Value="-1" />


    <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
        <ContentTemplate>
            <asp:HiddenField ID="PeriodoAddHid" runat="server" Value="0" />
            <asp:HiddenField ID="CebesAddHid" runat="server" Value="" />
            <asp:HiddenField ID="CecosAddHid" runat="server" Value="" />
            <asp:HiddenField ID="CuentasSAPAddHid" runat="server" Value="" />
            <asp:HiddenField ID="ArticulosGrupoAddHid" runat="server" Value="" />
            <asp:HiddenField ID="MaterialesAddHid" runat="server" Value="" />

            <asp:HiddenField ID="CebeSelHid" runat="server" />
            <asp:HiddenField ID="CecoSelHis" runat="server" />
            <asp:HiddenField ID="CSapSelHid" runat="server" />
            <asp:HiddenField ID="AGrupoSelHid" runat="server" />

            <asp:Button ID="AgregarCebeBtn" runat="server" OnClick="AgregarCebeBtn_Click" Style="display: none;" />
            <asp:Button ID="AgregarCecoBtn" runat="server" OnClick="AgregarCecoBtn_Click" Style="display: none;" />
            <asp:Button ID="AgregarCSapBtn" runat="server" OnClick="AgregarCSapBtn_Click" Style="display: none;" />
            <asp:Button ID="AgregarGArticuloBtn" runat="server" OnClick="AgregarGArticuloBtn_Click" Style="display: none;" />
            <asp:Button ID="AgregarMaterialBtn" runat="server" OnClick="AgregarMaterialBtn_Click" Style="display: none;" />

            <asp:HiddenField ID="CodigoIniEliminarHid" runat="server" Value="" />
            <asp:HiddenField ID="VersionEliminarHid" runat="server" Value="" />
            <asp:HiddenField ID="OrigenEliminarHid" runat="server" Value="" />

            <asp:Button ID="EliminarItemBtn" runat="server" OnClick="EliminarItemBtn_Click" Style="display: none;" />

            <asp:Button ID="EliminarPresupuestoBtn" runat="server" OnClick="EliminarPresupuestoBtn_Click" Style="display: none;" />
        </ContentTemplate>
    </asp:UpdatePanel>

    <script type="text/javascript">

        $(document).keypress(function (e) {
            if (e.which == 13) {
                return false;
            }

        });

        function OnInit(s, e) {
            AdjustSize();
        }

        function OnEndCallback(s, e) {
            AdjustSize();
        }

        function OnToolbarItemClick(s, e) {

            if (e.item.name == "ExportToXLSXH" || e.item.name == "ExportToXLSXV" || e.item.name == "ExportToXLSXPC") {
                e.processOnServer = true;
                e.usePostBack = true;
            }
            else if (e.item.name == "AddCebe") {
                AgregarItem(99, "");
                ASPxClientUtils.PreventEvent(e.htmlEvent);
            }
            else if (e.item.name == "Refresh") {
                presupuestotreelist.PerformCallback();
            }
            else if (e.item.name == "AddExcel") {
                DocumentoUploadSub();
                ASPxClientUtils.PreventEvent(e.htmlEvent);
            }
            else if (e.item.name == "DeletePre") {
                alertify.confirm("Eliminar Presupuesto", "¿Esta seguro que desea eliminar todo el presupuesto?<br><br>Se eliminará todo el presupuesto de las cuentas que usted puede presupuestar.",
                    function () {
                        $("#<%=EliminarPresupuestoBtn.ClientID%>").click();
                    },
                    function () {
                        //alertify.error('Cancel');
                    }).set('labels', { ok: 'Aceptar', cancel: 'Cancelar' });

                ASPxClientUtils.PreventEvent(e.htmlEvent);
            }
            else if (e.item.name == "CargaPre") {
                VerCargaSub();
                ASPxClientUtils.PreventEvent(e.htmlEvent);
            }



        }


        function VerCargaSub() {
            var url = "CargasConsultaSub.aspx?PresupuestoVersionId=" + versionddl.GetValue();;
            var height = $(window).height() - 60;
            var width = $(window).width() - 60;

            openShadowbox(url, "iframe", "", height, width);


        }

        function DocumentoUploadSub(proyectoId) {
            var url = "CargaPresupuestoSub.aspx?Version=" + versionddl.GetValue();;

            var height = 480;
            var width = 400;
            openShadowbox(url, "iframe", "", height, width);
        }



        function OnControlsInitialized(s, e) {
            ASPxClientUtils.AttachEventToElement(window, "resize", function (evt) {
                AdjustSize();
            });
        }

        function AdjustSize() {
            $("#<%=VersionInfoHid.ClientID%>").val(versionddl.GetValue());
            var height = Math.max(0, document.documentElement.clientHeight) - 200;
            presupuestotreelist.SetHeight(height);
        }

        function Actualizar() {
            $("#<%=VersionInfoHid.ClientID%>").val(versionddl.GetValue());
            presupuestotreelist.PerformCallback();
        }

        function ActualizarCarga(cantidad) {
            $("#<%=VersionInfoHid.ClientID%>").val(versionddl.GetValue());
            presupuestotreelist.PerformCallback();

            if (cantidad > 0) {
                alertify.alert("Importación Incompleta", "No se pudieron migrar " + cantidad + " filas, puede ver el detalle en la opción <b>Ver Cargas Realizadas</b>");
            }
        }

        function ActualizarPeriodo() {
            versionddl.PerformCallback();
        }

        function EditarPre(origen, codigo) {

            var url = "PresupuestoVersionSub.aspx?Origen=" + origen + "&Codigo=" + codigo + "&Version=" + versionddl.GetValue();;
            var height = 650;
            var width = 900;

            openShadowbox(url, "iframe", "", height, width);

            return false;
        }

        function AgregarItem(origen, codigoIni) {

            var ids = codigoIni.split(';');
            var url = "";

            if (origen == 0) {
                $("#<%=CebeSelHid.ClientID%>").val(ids[0]);
                url = "CECOBusquedaMulti.aspx?CEBECodigo=" + ids[0] + "&Anio=" + periodotxt.GetText();
            }
            else if (origen == 1) {
                $("#<%=CebeSelHid.ClientID%>").val(ids[0]);
                $("#<%=CecoSelHis.ClientID%>").val(ids[1]);
                url = "CuentaSAPBusquedaMulti.aspx?CEBECodigo=" + ids[0] + "&CECOCodigo=" + ids[1] + "&Anio=" + periodotxt.GetText() + "&ImpactaId=1";
            }
            else if (origen == 2) {
                $("#<%=CebeSelHid.ClientID%>").val(ids[0]);
                $("#<%=CecoSelHis.ClientID%>").val(ids[1]);
                $("#<%=CSapSelHid.ClientID%>").val(ids[2]);
                url = "ArticuloGrupoBusquedaMulti.aspx?CEBECodigo=" + ids[0] + "&CECOCodigo=" + ids[1] + "&CuentaSAPCodigo=" + ids[2] + "&Anio=" + periodotxt.GetText();
            }
            else if (origen == 3) {
                $("#<%=CebeSelHid.ClientID%>").val(ids[0]);
                $("#<%=CecoSelHis.ClientID%>").val(ids[1]);
                $("#<%=CSapSelHid.ClientID%>").val(ids[2]);
                $("#<%=AGrupoSelHid.ClientID%>").val(ids[3]);

                if (ids[3].substring(0, 1) != "S") {
                    url = "MaterialBusquedaMulti.aspx?CEBECodigo=" + ids[0] + "&CECOCodigo=" + ids[1] + "&CuentaSAPCodigo=" + ids[2] + "&ArticuloGrupoCodigo=" + ids[3] + "&Anio=" + periodotxt.GetText();
                } else {
                    url = "ContratoBusquedaMulti.aspx?CEBECodigo=" + ids[0] + "&CECOCodigo=" + ids[1] + "&CuentaSAPCodigo=" + ids[2] + "&ArticuloGrupoCodigo=" + ids[3] + "&Anio=" + periodotxt.GetText();
                }
            }
            else if (origen == 99) {
                url = "CEBEBusquedaMulti.aspx?Anio=" + periodotxt.GetText();
            }

            var height = $(window).height() - 60;
            var width = $(window).width() - 60;

            openShadowbox(url, "iframe", "", height, width);

            return false;

        }


        function CEBEMRel(ids) {
            if (ids != "") {
                $("#<%=PeriodoAddHid.ClientID%>").val(periodotxt.GetText());
                $("#<%=CebesAddHid.ClientID%>").val(ids);
                $("#<%=AgregarCebeBtn.ClientID%>").click();
            }
        }

        function CECOMRel(ids) {
            if (ids != "") {
                $("#<%=PeriodoAddHid.ClientID%>").val(periodotxt.GetText());
                $("#<%=CecosAddHid.ClientID%>").val(ids);
                $("#<%=AgregarCecoBtn.ClientID%>").click();
            }
        }

        function CuentaSAPMRel(ids) {
            if (ids != "") {
                $("#<%=PeriodoAddHid.ClientID%>").val(periodotxt.GetText());
                $("#<%=CuentasSAPAddHid.ClientID%>").val(ids);
                $("#<%=AgregarCSapBtn.ClientID%>").click();
            }
        }

        function ArticuloGrupoMRel(ids) {
            if (ids != "") {
                $("#<%=PeriodoAddHid.ClientID%>").val(periodotxt.GetText());
                $("#<%=ArticulosGrupoAddHid.ClientID%>").val(ids);
                $("#<%=AgregarGArticuloBtn.ClientID%>").click();
            }
        }

        function MaterialMRel(ids) {
            if (ids != "") {
                $("#<%=PeriodoAddHid.ClientID%>").val(periodotxt.GetText());
                $("#<%=MaterialesAddHid.ClientID%>").val(ids);
                $("#<%=AgregarMaterialBtn.ClientID%>").click();
            }
        }

        function ContratoMRel(ids) {
            if (ids != "") {
                $("#<%=PeriodoAddHid.ClientID%>").val(periodotxt.GetText());
                $("#<%=MaterialesAddHid.ClientID%>").val(ids);
                $("#<%=AgregarMaterialBtn.ClientID%>").click();
            }
        }



        function EliminarItem(codigoIni, nombre, origen) {

            alertify.confirm("Eliminar", "¿Esta seguro que desea eliminar<br>" + nombre + "?<br><br>Se borrará también la jerarquía inferior.",
                function () {
                    $("#<%=PeriodoAddHid.ClientID%>").val(periodotxt.GetText());

                    $("#<%=CodigoIniEliminarHid.ClientID%>").val(codigoIni);
                    $("#<%=VersionEliminarHid.ClientID%>").val(versionddl.GetValue());
                    $("#<%=OrigenEliminarHid.ClientID%>").val(origen);
                    $("#<%=EliminarItemBtn.ClientID%>").click();
                },
                function () {
                    //alertify.error('Cancel');
                }).set('labels', { ok: 'Aceptar', cancel: 'Cancelar' });

        }

    </script>
</asp:Content>
