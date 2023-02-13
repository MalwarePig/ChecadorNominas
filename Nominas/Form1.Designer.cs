namespace Nominas
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tblChecadaBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.checadorDataSet = new Nominas.Main.ChecadorDataSet();
            this.tblChecadaTableAdapter = new Nominas.Main.ChecadorDataSetTableAdapters.tblChecadaTableAdapter();
            this.tableAdapterManager = new Nominas.Main.ChecadorDataSetTableAdapters.TableAdapterManager();
            this.rdbBravo = new MaterialSkin.Controls.MaterialRadioButton();
            this.rdbMorelos = new MaterialSkin.Controls.MaterialRadioButton();
            this.tblTrabTurnoBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tblTrabTurnoTableAdapter = new Nominas.Main.ChecadorDataSetTableAdapters.tblTrabTurnoTableAdapter();
            this.Pruebas = new iTalk.iTalk_Button_1();
            this.dateFinal = new System.Windows.Forms.DateTimePicker();
            this.dateInicio = new System.Windows.Forms.DateTimePicker();
            this.iTalk_Panel2 = new iTalk.iTalk_Panel();
            this.gridSemanaTrabajador = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.tblChecadaBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.checadorDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblTrabTurnoBindingSource)).BeginInit();
            this.iTalk_Panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridSemanaTrabajador)).BeginInit();
            this.SuspendLayout();
            // 
            // tblChecadaBindingSource
            // 
            this.tblChecadaBindingSource.DataMember = "tblChecada";
            this.tblChecadaBindingSource.DataSource = this.checadorDataSet;
            // 
            // checadorDataSet
            // 
            this.checadorDataSet.DataSetName = "ChecadorDataSet";
            this.checadorDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // tblChecadaTableAdapter
            // 
            this.tblChecadaTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.tblAccesoProduccionTableAdapter = null;
            this.tableAdapterManager.tblAccesoTableAdapter = null;
            this.tableAdapterManager.tblAccionTableAdapter = null;
            this.tableAdapterManager.tblAreaTableAdapter = null;
            this.tableAdapterManager.tblAuditoriaTableAdapter = null;
            this.tableAdapterManager.tblCampoTableAdapter = null;
            this.tableAdapterManager.tblCategoriaCCTableAdapter = null;
            this.tableAdapterManager.tblCategoriaTableAdapter = null;
            this.tableAdapterManager.tblCeldaTableAdapter = null;
            this.tableAdapterManager.tblCentroDeTrabajoTableAdapter = null;
            this.tableAdapterManager.tblCentroTableAdapter = null;
            this.tableAdapterManager.tblChecadaOrdenServicioTableAdapter = null;
            this.tableAdapterManager.tblChecadaTableAdapter = this.tblChecadaTableAdapter;
            this.tableAdapterManager.tblChFotoRechTableAdapter = null;
            this.tableAdapterManager.tblChFotoTableAdapter = null;
            this.tableAdapterManager.tblChRechazadasTableAdapter = null;
            this.tableAdapterManager.tblColumnaReporteDinamicoTableAdapter = null;
            this.tableAdapterManager.tblComedorCostoHorarioTableAdapter = null;
            this.tableAdapterManager.tblComedorMenuTableAdapter = null;
            this.tableAdapterManager.tblComedorTableAdapter = null;
            this.tableAdapterManager.tblConceptoTableAdapter = null;
            this.tableAdapterManager.tblCorteCCTableAdapter = null;
            this.tableAdapterManager.tblDeptoTableAdapter = null;
            this.tableAdapterManager.tblEmpresaTableAdapter = null;
            this.tableAdapterManager.tblEstadoTableAdapter = null;
            this.tableAdapterManager.tblEventoTableAdapter = null;
            this.tableAdapterManager.tblFestivoTableAdapter = null;
            this.tableAdapterManager.tblFiltroFavoritoDetalleTableAdapter = null;
            this.tableAdapterManager.tblFiltroFavoritoTableAdapter = null;
            this.tableAdapterManager.tblFirmaTableAdapter = null;
            this.tableAdapterManager.tblFormatoReporteDinamicoTableAdapter = null;
            this.tableAdapterManager.tblGrupoTableAdapter = null;
            this.tableAdapterManager.tblHorarioAccesoTableAdapter = null;
            this.tableAdapterManager.tblHorarioComidaTableAdapter = null;
            this.tableAdapterManager.tblIncidenciaOutNomTableAdapter = null;
            this.tableAdapterManager.tblIncidenciasHTableAdapter = null;
            this.tableAdapterManager.tblIncidenciasTableAdapter = null;
            this.tableAdapterManager.tblLectorTableAdapter = null;
            this.tableAdapterManager.tblLineaCCTableAdapter = null;
            this.tableAdapterManager.tblLocalidadTableAdapter = null;
            this.tableAdapterManager.tblLogTableAdapter = null;
            this.tableAdapterManager.tblMensajeTableAdapter = null;
            this.tableAdapterManager.tblMenuTableAdapter = null;
            this.tableAdapterManager.tblMetodoChecadoTableAdapter = null;
            this.tableAdapterManager.tblOpcionesTableAdapter = null;
            this.tableAdapterManager.tblOrdenTableAdapter = null;
            this.tableAdapterManager.tblParametrosTableAdapter = null;
            this.tableAdapterManager.tblPermisoTableAdapter = null;
            this.tableAdapterManager.tblPlantillaTableAdapter = null;
            this.tableAdapterManager.tblPlanTurnoDiarioTableAdapter = null;
            this.tableAdapterManager.tblPlanTurnoTableAdapter = null;
            this.tableAdapterManager.tblProyectoCCTableAdapter = null;
            this.tableAdapterManager.tblRechazoTableAdapter = null;
            this.tableAdapterManager.tblReglaPlanTurnoTableAdapter = null;
            this.tableAdapterManager.tblReporteDinamicoTableAdapter = null;
            this.tableAdapterManager.tblRotacionTableAdapter = null;
            this.tableAdapterManager.tblRotaTurnoTableAdapter = null;
            this.tableAdapterManager.tblSincroHoraTableAdapter = null;
            this.tableAdapterManager.tblSupervisorTableAdapter = null;
            this.tableAdapterManager.tblSupHuella00bTableAdapter = null;
            this.tableAdapterManager.tblSupHuella00TableAdapter = null;
            this.tableAdapterManager.tblSupHuella03TableAdapter = null;
            this.tableAdapterManager.tblSupHuellaTableAdapter = null;
            this.tableAdapterManager.tblTarjetaTableAdapter = null;
            this.tableAdapterManager.tblTerminal_tblTrabajadorTableAdapter = null;
            this.tableAdapterManager.tblTerminalLectorTableAdapter = null;
            this.tableAdapterManager.tblTerminalRelevadorTableAdapter = null;
            this.tableAdapterManager.tblTerminalTableAdapter = null;
            this.tableAdapterManager.tblTiempoExtraAutTableAdapter = null;
            this.tableAdapterManager.tblTimbreTableAdapter = null;
            this.tableAdapterManager.tblTipoChecadaTableAdapter = null;
            this.tableAdapterManager.tblTipoDeCambioTableAdapter = null;
            this.tableAdapterManager.tblTipoEmpleadoTableAdapter = null;
            this.tableAdapterManager.tblTipoPlantillaTableAdapter = null;
            this.tableAdapterManager.tblTmpAltaWSTableAdapter = null;
            this.tableAdapterManager.tblTmpBajaWSTableAdapter = null;
            this.tableAdapterManager.tblTmpCambioTrabajadorWSTableAdapter = null;
            this.tableAdapterManager.tblTmpCheEvaTableAdapter = null;
            this.tableAdapterManager.tblTmpCheTableAdapter = null;
            this.tableAdapterManager.tblTmpDepCenTableAdapter = null;
            this.tableAdapterManager.tblTmpInicioRotacionTableAdapter = null;
            this.tableAdapterManager.tblTmpResponsesWSTableAdapter = null;
            this.tableAdapterManager.tblTmpRotacionTableAdapter = null;
            this.tableAdapterManager.tblTmpTrabajador_InicioRotacionTableAdapter = null;
            this.tableAdapterManager.tblTmpTransferTrabajadorWSTableAdapter = null;
            this.tableAdapterManager.tblTmpTraTableAdapter = null;
            this.tableAdapterManager.tblTmpTraTurTableAdapter = null;
            this.tableAdapterManager.tblTmpTurTableAdapter = null;
            this.tableAdapterManager.tblTrabAccesoTableAdapter = null;
            this.tableAdapterManager.tblTrabajador_tblCentroDeTrabajoTableAdapter = null;
            this.tableAdapterManager.tblTrabajadorTableAdapter = null;
            this.tableAdapterManager.tblTrabAuditarTableAdapter = null;
            this.tableAdapterManager.tblTrabCampoTableAdapter = null;
            this.tableAdapterManager.tblTrabCara06TableAdapter = null;
            this.tableAdapterManager.tblTrabCCTableAdapter = null;
            this.tableAdapterManager.tblTrabConceptoTableAdapter = null;
            this.tableAdapterManager.tblTrabHanvonTableAdapter = null;
            this.tableAdapterManager.tblTrabHuella00bTableAdapter = null;
            this.tableAdapterManager.tblTrabHuella00TableAdapter = null;
            this.tableAdapterManager.tblTrabHuella01sTableAdapter = null;
            this.tableAdapterManager.tblTrabHuella01TableAdapter = null;
            this.tableAdapterManager.tblTrabHuella02TableAdapter = null;
            this.tableAdapterManager.tblTrabHuella03TableAdapter = null;
            this.tableAdapterManager.tblTrabHuella04TableAdapter = null;
            this.tableAdapterManager.tblTrabHuella05TableAdapter = null;
            this.tableAdapterManager.tblTrabHuella06TableAdapter = null;
            this.tableAdapterManager.tblTrabHuellaTableAdapter = null;
            this.tableAdapterManager.tblTrabIncidAutTableAdapter = null;
            this.tableAdapterManager.tblTrabLectorRelevadorTableAdapter = null;
            this.tableAdapterManager.tblTrabManoTableAdapter = null;
            this.tableAdapterManager.tblTrabMensaje02TableAdapter = null;
            this.tableAdapterManager.tblTrabMensajeTableAdapter = null;
            this.tableAdapterManager.tblTrabReglaTurnoTableAdapter = null;
            this.tableAdapterManager.tblTrabSalarioTableAdapter = null;
            this.tableAdapterManager.tblTrabTarjeta06TableAdapter = null;
            this.tableAdapterManager.tblTrabTurnoTableAdapter = null;
            this.tableAdapterManager.tblTrabUsuario06TableAdapter = null;
            this.tableAdapterManager.tblTurnoTableAdapter = null;
            this.tableAdapterManager.tblUDPAccesoTableAdapter = null;
            this.tableAdapterManager.tblUDPActivoTableAdapter = null;
            this.tableAdapterManager.tblUDPChecadaTableAdapter = null;
            this.tableAdapterManager.tblUsuario_tblCentroDeTrabajoTableAdapter = null;
            this.tableAdapterManager.tblUsuario_tblEmpresaTableAdapter = null;
            this.tableAdapterManager.tblUsuarioCeldaTableAdapter = null;
            this.tableAdapterManager.tblUsuarioDeptoTableAdapter = null;
            this.tableAdapterManager.tblUsuarioMenuTableAdapter = null;
            this.tableAdapterManager.tblUsuarioTableAdapter = null;
            this.tableAdapterManager.tblWFS_SwipeResultTableAdapter = null;
            this.tableAdapterManager.tblWFSPayCodeTableAdapter = null;
            this.tableAdapterManager.tblWFSTransactionTypeTableAdapter = null;
            this.tableAdapterManager.tblZonaHorariaTableAdapter = null;
            this.tableAdapterManager.UpdateOrder = Nominas.Main.ChecadorDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // rdbBravo
            // 
            this.rdbBravo.AutoSize = true;
            this.rdbBravo.Checked = true;
            this.rdbBravo.Depth = 0;
            this.rdbBravo.Font = new System.Drawing.Font("Roboto", 10F);
            this.rdbBravo.Location = new System.Drawing.Point(154, 67);
            this.rdbBravo.Margin = new System.Windows.Forms.Padding(0);
            this.rdbBravo.MouseLocation = new System.Drawing.Point(-1, -1);
            this.rdbBravo.MouseState = MaterialSkin.MouseState.HOVER;
            this.rdbBravo.Name = "rdbBravo";
            this.rdbBravo.Ripple = true;
            this.rdbBravo.Size = new System.Drawing.Size(64, 30);
            this.rdbBravo.TabIndex = 4;
            this.rdbBravo.TabStop = true;
            this.rdbBravo.Text = "Bravo";
            this.rdbBravo.UseVisualStyleBackColor = true;
            // 
            // rdbMorelos
            // 
            this.rdbMorelos.AutoSize = true;
            this.rdbMorelos.Depth = 0;
            this.rdbMorelos.Font = new System.Drawing.Font("Roboto", 10F);
            this.rdbMorelos.Location = new System.Drawing.Point(229, 199);
            this.rdbMorelos.Margin = new System.Windows.Forms.Padding(0);
            this.rdbMorelos.MouseLocation = new System.Drawing.Point(-1, -1);
            this.rdbMorelos.MouseState = MaterialSkin.MouseState.HOVER;
            this.rdbMorelos.Name = "rdbMorelos";
            this.rdbMorelos.Ripple = true;
            this.rdbMorelos.Size = new System.Drawing.Size(79, 30);
            this.rdbMorelos.TabIndex = 5;
            this.rdbMorelos.Text = "Morelos";
            this.rdbMorelos.UseVisualStyleBackColor = true;
            // 
            // tblTrabTurnoBindingSource
            // 
            this.tblTrabTurnoBindingSource.DataMember = "tblTrabTurno";
            this.tblTrabTurnoBindingSource.DataSource = this.checadorDataSet;
            // 
            // tblTrabTurnoTableAdapter
            // 
            this.tblTrabTurnoTableAdapter.ClearBeforeFill = true;
            // 
            // Pruebas
            // 
            this.Pruebas.BackColor = System.Drawing.Color.Transparent;
            this.Pruebas.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.Pruebas.Image = null;
            this.Pruebas.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Pruebas.Location = new System.Drawing.Point(345, 171);
            this.Pruebas.Name = "Pruebas";
            this.Pruebas.Size = new System.Drawing.Size(166, 40);
            this.Pruebas.TabIndex = 1;
            this.Pruebas.Text = "Importar BD";
            this.Pruebas.TextAlignment = System.Drawing.StringAlignment.Center;
            this.Pruebas.Click += new System.EventHandler(this.Pruebas_Click);
            // 
            // dateFinal
            // 
            this.dateFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateFinal.Location = new System.Drawing.Point(288, 93);
            this.dateFinal.Name = "dateFinal";
            this.dateFinal.Size = new System.Drawing.Size(200, 23);
            this.dateFinal.TabIndex = 12;
            this.dateFinal.Value = new System.DateTime(2019, 6, 25, 0, 0, 0, 0);
            // 
            // dateInicio
            // 
            this.dateInicio.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateInicio.Location = new System.Drawing.Point(288, 67);
            this.dateInicio.Name = "dateInicio";
            this.dateInicio.Size = new System.Drawing.Size(200, 23);
            this.dateInicio.TabIndex = 11;
            this.dateInicio.Value = new System.DateTime(2019, 6, 25, 0, 0, 0, 0);
            // 
            // iTalk_Panel2
            // 
            this.iTalk_Panel2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.iTalk_Panel2.BackColor = System.Drawing.Color.Transparent;
            this.iTalk_Panel2.Controls.Add(this.gridSemanaTrabajador);
            this.iTalk_Panel2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.iTalk_Panel2.Location = new System.Drawing.Point(16, 249);
            this.iTalk_Panel2.Name = "iTalk_Panel2";
            this.iTalk_Panel2.Padding = new System.Windows.Forms.Padding(5);
            this.iTalk_Panel2.Size = new System.Drawing.Size(721, 448);
            this.iTalk_Panel2.TabIndex = 16;
            this.iTalk_Panel2.Text = "iTalk_Panel2";
            // 
            // gridSemanaTrabajador
            // 
            this.gridSemanaTrabajador.AllowUserToAddRows = false;
            this.gridSemanaTrabajador.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gridSemanaTrabajador.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridSemanaTrabajador.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridSemanaTrabajador.Location = new System.Drawing.Point(5, 5);
            this.gridSemanaTrabajador.Name = "gridSemanaTrabajador";
            this.gridSemanaTrabajador.Size = new System.Drawing.Size(711, 438);
            this.gridSemanaTrabajador.TabIndex = 2;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(749, 698);
            this.Controls.Add(this.iTalk_Panel2);
            this.Controls.Add(this.dateFinal);
            this.Controls.Add(this.dateInicio);
            this.Controls.Add(this.rdbMorelos);
            this.Controls.Add(this.rdbBravo);
            this.Controls.Add(this.Pruebas);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load_1);
            ((System.ComponentModel.ISupportInitialize)(this.tblChecadaBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.checadorDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblTrabTurnoBindingSource)).EndInit();
            this.iTalk_Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridSemanaTrabajador)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private iTalk.iTalk_Button_1 Pruebas;
        private Main.ChecadorDataSet checadorDataSet;
        private System.Windows.Forms.BindingSource tblChecadaBindingSource;
        private Main.ChecadorDataSetTableAdapters.tblChecadaTableAdapter tblChecadaTableAdapter;
        private Main.ChecadorDataSetTableAdapters.TableAdapterManager tableAdapterManager;
        private MaterialSkin.Controls.MaterialRadioButton rdbBravo;
        private MaterialSkin.Controls.MaterialRadioButton rdbMorelos;
        private System.Windows.Forms.BindingSource tblTrabTurnoBindingSource;
        private Main.ChecadorDataSetTableAdapters.tblTrabTurnoTableAdapter tblTrabTurnoTableAdapter;
        private System.Windows.Forms.DateTimePicker dateFinal;
        private System.Windows.Forms.DateTimePicker dateInicio;
        private iTalk.iTalk_Panel iTalk_Panel2;
        private System.Windows.Forms.DataGridView gridSemanaTrabajador;
    }
}