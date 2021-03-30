using Amazon;
using Amazon.SecretsManager;
using Amazon.SecretsManager.Model;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

public static class Utility
{
    static AmazonSecretsManagerClient clientSecretManager = new AmazonSecretsManagerClient(RegionEndpoint.APSoutheast1);
    static string conn = null;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string Env(string key)
    {
        return Environment.GetEnvironmentVariable(key);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static Task<GetSecretValueResponse> SecretManager(string key)
    {
        
        var task = clientSecretManager.GetSecretValueAsync(new GetSecretValueRequest
        {
            SecretId = key
        });

        return task;
    }


    public static async Task<string> ConnectionString()
    {
        if(conn == null)
        {
            var user = string.Empty;
            var pass = string.Empty;
            
            if (!string.IsNullOrEmpty(Env("SECRET_NAME")))
            {
                var json = await SecretManager(Env("SECRET_NAME"));
                var data = JsonDocument.Parse(json.SecretString).RootElement;
                user = data.GetProperty("pg_username").GetString();
                pass = data.GetProperty("pg_password").GetString();
            }
            else
            {
                user = Env("USER");
                pass = Env("PASSWORD");
            }
           

            conn = $"Host={Env("HOST")};PORT={Env("PORT")};Database={Env("DATABASE")};Username={user};Password={pass};";

        }

        return conn;
    }

    public static async Task<NpgsqlConnection> CreateConnection()
    {
        return new NpgsqlConnection(await ConnectionString());
    }

    public static async Task<DbDataReader> ExecuteReaderAsync(this NpgsqlConnection connection, string sql)
    {
        await using (var cmd = connection.CreateCommand())
        {
            cmd.CommandText = sql;
            if (connection.State == System.Data.ConnectionState.Closed)
            {
                await connection.OpenAsync();
            }
            return await cmd.ExecuteReaderAsync(System.Data.CommandBehavior.CloseConnection);
        }
    }

    public static DbDataReader ExecuteReader(this NpgsqlConnection connection, string sql)
    {
        using (var cmd = connection.CreateCommand())
        {
            cmd.CommandText = sql;
            if (connection.State == System.Data.ConnectionState.Closed)
            {
                connection.Open();
            }
            return cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
        }
    }

    public static async Task<int> ExecuteNonQueryAsync(this NpgsqlConnection connection, string sql)
    {
        var result = 0;
        using (DbCommand cmd = connection.CreateCommand())
        {
            cmd.CommandText = sql;
            if (connection.State == System.Data.ConnectionState.Closed)
            {
                await connection.OpenAsync();
            }
            result = await cmd.ExecuteNonQueryAsync();
            connection.Close();
        }
        return result;
    }

    public static int ExecuteNonQuery(this NpgsqlConnection connection, string sql)
    {
        var result = 0;
        using (DbCommand cmd = connection.CreateCommand())
        {
            cmd.CommandText = sql;
            if (connection.State == System.Data.ConnectionState.Closed)
            {
                connection.Open();
            }
            result = cmd.ExecuteNonQuery();
            connection.Close();
        }
        return result;
    }

    public static async Task<object> ExecuteScalarAsync(this NpgsqlConnection connection, string sql)
    {
        object result;
        using (DbCommand cmd = connection.CreateCommand())
        {
            cmd.CommandText = sql;

            if (connection.State == System.Data.ConnectionState.Closed)
            {
                await connection.OpenAsync();
            }
            result = await cmd.ExecuteScalarAsync().ConfigureAwait(false);
            connection.Close();
        }
        return result;
    }

    public static object ExecuteScalar(this NpgsqlConnection connection, string sql)
    {
        object result;
        using (DbCommand cmd = connection.CreateCommand())
        {
            cmd.CommandText = sql;

            if (connection.State == System.Data.ConnectionState.Closed)
            {
                connection.Open();
            }
            result =  cmd.ExecuteScalar();
            connection.Close();
        }
        return result;
    }
}

