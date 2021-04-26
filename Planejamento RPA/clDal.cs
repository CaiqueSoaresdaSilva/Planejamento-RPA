using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Data.Common;
using System.Text;

public class ParameterBuilder : IDisposable
{

    private string _ParamName;
    private object _Value;
    private DbType _DbType = System.Data.DbType.String;
    private ParameterDirection _Direction = ParameterDirection.Input;
    private int _Size = 0;
    private string _SourceColumn = string.Empty;
    private DataRowVersion _SourceVersion = DataRowVersion.Current;

    private bool _SourceColumnNullMapping = false;

    public enum TipoExecucao
    {
        ExecutaQuery = 0,
        RetornaParam = 1,
        RetornaTabela = 2
    }


    public ParameterBuilder(string ParameterName, DbType dbType, object Value, ParameterDirection Direction, int Size)
    {
        this._ParamName = ParameterName;
        this._DbType = dbType;
        this._Value = Value;
        this._Direction = Direction;
        this._Size = Size;
    }


    public string ParamName
    {
        get { return this._ParamName; }
        set { this._ParamName = value; }
    }

    public object Value
    {
        get { return _Value; }
        set { _Value = value; }
    }

    public DbType DbType
    {
        get { return _DbType; }
        set { _DbType = value; }
    }

    public ParameterDirection Direction
    {
        get { return _Direction; }
        set { _Direction = value; }
    }

    public int Size
    {
        get { return this._Size; }
        set { this._Size = value; }
    }


    #region "IDisposable Support"

    private bool disposedValue;
    protected virtual void Dispose(bool disposing)
    {
        if (!this.disposedValue)
        {
            if (disposing)
            {
                _ParamName = null;
                _Value = null;
                _DbType = 0;
                _Direction = 0;
                _Size = 0;
                _SourceColumn = null;
                _SourceVersion = 0;
                //   _SourceColumnNullMapping = Boolean.;
            }
        }
        this.disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

}

public class clDALSQL : IDisposable
{

    System.Data.SqlClient.SqlConnection Conn = new System.Data.SqlClient.SqlConnection();
    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand();
    System.Data.SqlClient.SqlDataAdapter Adapt = new System.Data.SqlClient.SqlDataAdapter();
    DataTable Dt;

    public List<SqlParameter> Parametros = new List<SqlParameter>();
    public enum AmbienteExecucao : int
    {
        Desenvolvimento = 0,
        Producao = 1,
        Homologacao = 2,
        DW = 3,
        BDE = 4
    }

    public enum InstanciaExecucao : int
    {
        Normal = 0,
        Alog = 1
    }


    private AmbienteExecucao mAmbienteExecucao = AmbienteExecucao.Producao;
    public AmbienteExecucao Ambiente
    {
        get { return mAmbienteExecucao; }
        set { mAmbienteExecucao = value; }
    }

    private string retorna_ip()
    {
        string slqIp;
        string diretorio = @"C:\IPdb.ini";

        System.IO.StreamReader file = new System.IO.StreamReader(diretorio);
        slqIp = file.ReadLine();

        return slqIp;

    }

    private string RetornaStrConexao(string NomeBanco, InstanciaExecucao Instancia = InstanciaExecucao.Normal)
    {
        string functionReturnValue = null;
        functionReturnValue = "";
        if (mAmbienteExecucao == AmbienteExecucao.Producao)
        {
            if (Instancia == InstanciaExecucao.Normal)
            {
                functionReturnValue = "Password=pr0dbdn07w1LL14mjAmA15;Persist Security Info=True;User ID=db_prod2;Initial Catalog=" + NomeBanco + ";Data Source=" + retorna_ip();
            }
            else if (Instancia == InstanciaExecucao.Alog)
            {
                functionReturnValue = "Password=!#m0d&rn0#!;Persist Security Info=True;User ID=backseg_sa;Initial Catalog=" + NomeBanco + ";Data Source=177.126.186.178";
            }
        }
        else if (mAmbienteExecucao == AmbienteExecucao.DW)
        {
            functionReturnValue = "Password=CHk2v3;Persist Security Info=True;User ID=db_prod;Initial Catalog=" + NomeBanco + ";Data Source=10.0.0.22";
        }
        else if (mAmbienteExecucao == AmbienteExecucao.BDE)
        {
            functionReturnValue = "Password=99Y34w8;Persist Security Info=True;User ID=excelcef;Initial Catalog=" + NomeBanco + ";Data Source=10.0.0.76";
        }
        else
        {
            functionReturnValue = "Password=!#m0d&rn0#!;Persist Security Info=True;User ID=backseg_sa;Initial Catalog=" + NomeBanco + ";Data Source=177.126.186.178";
            //RetornaStrConexao = "Password=Wr3327;Persist Security Info=True;User ID=db_dev;Initial Catalog=" & NomeBanco & ";Data Source=200.219.199.122"
        }
        return functionReturnValue;

    }

    private string RetornaStrConexao_PostGre(string NomeBanco)
    {
        string functionReturnValue = null;
        if (mAmbienteExecucao == AmbienteExecucao.Producao)
        {
            functionReturnValue = "Server=200.213.221.131;Port=5432;UserId=integracao;Password=;Database=" + NomeBanco;
        }
        else
        {
            functionReturnValue = "Server=200.213.221.131;Port=5432;UserId=integracao;Password=;Database=" + NomeBanco;
        }
        return functionReturnValue;
    }

    public void ClearParameters()
    {
        Parametros.Clear();
    }

    public void ExecutaProcedure(string NomeProcedure, string NomeBanco, int ExecutionTimeOut = 0, InstanciaExecucao Instancia = InstanciaExecucao.Normal)
    {
        SqlParameter obParam = default(SqlParameter);

        if (Conn.State == ConnectionState.Open)
        {
            Conn.Close();
        }
        Conn.ConnectionString = RetornaStrConexao(NomeBanco, Instancia);
        Conn.Open();
        Command.Connection = Conn;
        Command.CommandText = NomeProcedure;
        Command.Parameters.Clear();
        Command.CommandType = CommandType.StoredProcedure;
        foreach (SqlParameter obParam_loopVariable2 in Parametros)
        {
            obParam = obParam_loopVariable2;
            Command.Parameters.Add(obParam);
        }
        if (ExecutionTimeOut > 0)
        {
            Command.CommandTimeout = ExecutionTimeOut;
        }
        else
        {
            Command.CommandTimeout = 0;
        }
        Command.ExecuteNonQuery();

        //APOS RETORNAR A CONSULTA ATUALIZAR OS PARAMETROS
        Parametros.Clear();

        foreach (SqlParameter obParam_loopVariable3 in Command.Parameters)
        {
            obParam = obParam_loopVariable3;
            Parametros.Add(obParam);
        }

        if (Conn.State == ConnectionState.Open)
        {
            Conn.Close();
        }

    }

    public void ExecutaQuery(string Query, string NomeBanco, int ExecutionTimeOut = 0, InstanciaExecucao Instancia = InstanciaExecucao.Normal)
    {

        if (Conn.State == ConnectionState.Open)
        {
            Conn.Close();
        }
        Conn.ConnectionString = RetornaStrConexao(NomeBanco, Instancia);
        Conn.Open();
        Command.Connection = Conn;
        Command.CommandText = Query;
        Command.Parameters.Clear();
        Command.CommandType = CommandType.Text;

        if (ExecutionTimeOut > 0)
        {
            Command.CommandTimeout = ExecutionTimeOut;
        }
        else
        {
            Command.CommandTimeout = 0;
        }
        Command.ExecuteNonQuery();

        //APOS RETORNAR A CONSULTA ATUALIZAR OS PARAMETROS
        Parametros.Clear();

        if (Conn.State == ConnectionState.Open)
        {
            Conn.Close();
        }

    }

    public void AddParameters(string Nome, object Valor, SqlDbType Type, ParameterDirection Direcao = ParameterDirection.Input, int Tamanho = 0)
    {
        SqlParameter obParam = default(SqlParameter);
        obParam = new SqlParameter(Nome, Valor);
        obParam.Direction = Direcao;
        if (Tamanho > 0)
        {
            obParam.Size = Tamanho;
        }
        else if (Tamanho == -1)
        {
            //    obParam.DbType = SqlDbType.VarChar;
            obParam.Size = -1;
        }
        else
        {
            obParam.Size = 0;
        }
        Parametros.Add(obParam);
    }

    public DataTable RetornaTabela(string NomeProcedure, string NomeBanco, int ExecutionTimeOut = 0, InstanciaExecucao Instancia = InstanciaExecucao.Normal)
    {
        DataTable functionReturnValue = default(DataTable);
        DataTable obDT = new DataTable();

        try
        {
            if (Conn.State == ConnectionState.Open)
            {
                Conn.Close();
            }
            Conn.ConnectionString = RetornaStrConexao(NomeBanco, Instancia);
            Conn.Open();
            Command.Connection = Conn;
            Command.CommandText = NomeProcedure;
            Command.Parameters.Clear();
            Command.CommandType = CommandType.StoredProcedure;
            foreach (Object obParam_loopVariable in Parametros)
            {
                // obParam = obParam_loopVariable;
                Command.Parameters.Add(obParam_loopVariable);
            }
            if (ExecutionTimeOut > 0)
            {
                Command.CommandTimeout = ExecutionTimeOut;
            }
            else
            {
                Command.CommandTimeout = 0;
            }
            Adapt.SelectCommand = Command;

            Adapt.Fill(obDT);

        }
        catch (Exception ex)
        {

        }
        finally
        {
            functionReturnValue = obDT;
        }
        return functionReturnValue;


    }


    #region "IDisposable Support"

    private bool disposedValue;

    protected virtual void Dispose(bool disposing)
    {
        if (!this.disposedValue)
        {
            if (disposing)
            {

            }
            if (Conn.State == ConnectionState.Open)
            {
                Conn.Close();
            }
            Conn.Dispose();
            Command.Dispose();
            Adapt.Dispose();
            Parametros.Clear();
            Parametros = null;
            Dt = null;

        }
        this.disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
    #endregion

}



