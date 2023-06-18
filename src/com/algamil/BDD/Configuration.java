


package com.algamil.BDD;

import static com.algamil.BDD.Parameters.*;

import java.sql.*;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author FaridMO
 */
public class Configuration {

    static Connection connection;
    static Statement statement;
    static PreparedStatement pst;
    static ResultSet rs;
    static String SQL;

    /* Connexion à la base de données */
    public static Connection dbConn() {
        
        try {
            Class.forName(JDBC_DRIVER);
            connection = DriverManager.getConnection(HOST_DB, USERNAME_DB, PASSWORD_DB);

            // System.out.println("connexion Réussie");
        } catch (ClassNotFoundException ex) {
            // System.out.println("echec");
           Logger.getLogger(Configuration.class.getName()).log(Level.SEVERE, null, ex);
           
        } catch (SQLException ex) {
            Logger.getLogger(Configuration.class.getName()).log(Level.SEVERE, null, ex);
        }
        return connection;
    }

    public static Connection closeConn() {
        try {
            connection.close();
        } catch (SQLException ex) {
            Logger.getLogger(Configuration.class.getName()).log(Level.SEVERE, null, ex);
        }
        return connection;
    }

    /* Execution des requêtes */
    public static ResultSet executionQuery(String sql) {
        dbConn();
        ResultSet resultSet = null;
        try {
            //statement = connection.createStatement();
            pst = dbConn().prepareStatement(sql);
            resultSet = pst.executeQuery(sql);
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return resultSet;
    }

    /* Mise à jour de l'execution */
    public static String executionUpdate(String sql) {

        String result = "";
        try {
            dbConn();

            statement = connection.createStatement();
            statement.executeUpdate(sql);
            result = sql;
            //System.out.println("réussie !");
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return result;
    }

    /* Insertion des requêtes */
    public static String insert(String table, String[] colonnes, String[] valeurs) {
        SQL = "INSERT INTO " + table + " (";

        for (int i = 0; i < colonnes.length; i++) {
            SQL += colonnes[i];

            if (i < colonnes.length - 1) {
                SQL += ",";
            }
        }
        SQL += ") VALUES (";

        for (int i = 0; i < valeurs.length; i++) {
            SQL += "'" + valeurs[i] + "'";

            if (i < valeurs.length - 1) {
                SQL += ",";
            }
        }

        SQL += ")";
        System.out.println(SQL);
        return executionUpdate(SQL);
    }

    /* Affichage des tables */
    public static ResultSet selectAll(String table) {
        dbConn();
        SQL = "SELECT * FROM " + table;
        System.out.println(SQL);
        return executionQuery(SQL);
    }

    public static ResultSet selectCondition(String table, String etat) {
        dbConn();
        SQL = "SELECT COUNT(*) FROM " + table + " WHERE num_cni=" + etat;
        System.out.println(SQL);
        return executionQuery(SQL);

    }
    
        public static ResultSet selectConditionLogin(String table, String etat) {
        dbConn();
        SQL = "SELECT * FROM " + table + " WHERE "+ etat;
        System.out.println(SQL);
        return executionQuery(SQL);

    }

    public static ResultSet searchByDepartment(String table, String colonnes, String etat) {
        SQL = "SELECT "+colonnes +" FROM " + table + " WHERE department LIKE '" + etat + "%'";
        return executionQuery(SQL);
    }

    public static ResultSet searchByHostname(String table, String colonnes, String etat) {
        SQL = "SELECT " +colonnes+ " FROM " + table + " WHERE hostname LIKE '" + etat + "%'";
        return executionQuery(SQL);
    }
    

    public static ResultSet select(String[] nomColonne, String table) {
        dbConn();
        SQL = "SELECT ";
        for (int i = 0; i <= nomColonne.length - 1; i++) {
            SQL += nomColonne[i];
            if (i < nomColonne.length - 1) {
                SQL += " ,";
            }
        }
        SQL += " FROM " + table;
        return executionQuery(SQL);
    }

    public static String updateTable(String table, String[] colonnes, String[] valeurs, String idB) {
        dbConn();

        SQL = "UPDATE " + table + " SET ";
        for (int i = 0; i < colonnes.length; i++) {
            SQL += colonnes[i] + "='" + valeurs[i] + "'";

            if (i < colonnes.length - 1) {
                SQL += ",";
            }
        }
        SQL += " WHERE pid=" + idB;

        System.out.println(SQL);
        return executionUpdate(SQL);
    }

    public static String deleteColumn(String table, String idB) {
        dbConn();
        SQL = "DELETE FROM " + table + " WHERE pid=" + idB;

        return executionUpdate(SQL);
    }
    
    public static void idReload(String table){
        dbConn();
        String SQL1 = "ALTER TABLE "+table+" DROP pid";
        executionUpdate(SQL1);
        String SQL2 = "ALTER TABLE "+table+" ADD pid int NOT NULL AUTO_INCREMENT PRIMARY KEY FIRST";
        executionUpdate(SQL2);
    }

}
