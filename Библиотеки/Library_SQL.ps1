
Class SQL_Query{
    
    $connectionString = “Provider=sqloledb; Data Source=example; Initial Catalog=example; Persist Security Info=False; User ID=example; Password=example;“


    [int]fCheckTableExist($tableName){

        $sql_check_table_exist = “DECLARE @A INT;
                                  IF (EXISTS (SELECT * 
                                                   FROM INFORMATION_SCHEMA.TABLES 
                                                   WHERE TABLE_NAME = '$tableName'))
                                  BEGIN
                                      SET @A = 1;
                                      SELECT @A as 'EXIST'
                                  END
                                  ELSE 
                                  BEGIN
	                                  SET @A = 0;
                                      SELECT @A as 'EXIST'
                                  END”

        $connection = New-Object System.Data.OleDb.OleDbConnection $this.connectionString
        $command = New-Object System.Data.OleDb.OleDbCommand $sql_check_table_exist,$connection
        $connection.Open()
        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $command
        $dataset = New-Object System.Data.DataSet
        [void] $adapter.Fill($dataSet)
        $connection.Close()
        $rows=($dataset.Tables | Select-Object -Expand Rows)
    
        Return $rows.EXIST
 
    }

    fCreateTable($tableName){
    
        $sql_create_table = "BEGIN TRANSACTION
                             SET QUOTED_IDENTIFIER ON
                             SET ARITHABORT ON
                             SET NUMERIC_ROUNDABORT OFF
                             SET CONCAT_NULL_YIELDS_NULL ON
                             SET ANSI_NULLS ON
                             SET ANSI_PADDING ON
                             SET ANSI_WARNINGS ON
                             COMMIT
                             BEGIN TRANSACTION
                             CREATE TABLE $tableName (
                                 date datetime,
                                 operation ntext,
                                 filename ntext);
                             COMMIT"

        $connection = New-Object System.Data.OleDb.OleDbConnection $this.connectionString
        $command = New-Object System.Data.OleDb.OleDbCommand $sql_create_table,$connection
        $connection.Open()
        $command = New-Object data.OleDb.OleDbCommand $sql_create_table
        $command.connection = $connection
        $rowsAffected = $command.ExecuteNonQuery()

    }

    fInsertLog($tableName, $date, $operation, $filename){
    
        $sql_insertLog = "insert into $tableName (date, operation, filename) Values ('$date', '$operation', '$filename')"
        $connection = New-Object System.Data.OleDb.OleDbConnection $this.connectionString
        $command = New-Object System.Data.OleDb.OleDbCommand $sql_insertLog,$connection
        $connection.Open()
        $command = New-Object data.OleDb.OleDbCommand $sql_insertLog
        $command.connection = $connection
        $rowsAffected = $command.ExecuteNonQuery()

    
    }

    fInsertErrorsLog($tableName, $date, $scripterrors){
    
        $sql_insertLog = "insert into $tableName (date, scripterrors) Values ('$date', '$scripterrors')"
        $connection = New-Object System.Data.OleDb.OleDbConnection $this.connectionString
        $command = New-Object System.Data.OleDb.OleDbCommand $sql_insertLog,$connection
        $connection.Open()
        $command = New-Object data.OleDb.OleDbCommand $sql_insertLog
        $command.connection = $connection
        $rowsAffected = $command.ExecuteNonQuery()

    
    }

    [object]fSelectFromTable($tableName){
    
        $sql_Select = "SELECT * FROM $tableName"

        $connection = New-Object System.Data.OleDb.OleDbConnection $this.connectionString
        $command = New-Object System.Data.OleDb.OleDbCommand $sql_Select,$connection
        $connection.Open()
        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $command
        $dataset = New-Object System.Data.DataSet
        [void] $adapter.Fill($dataSet)
        $connection.Close()
        $rows=($dataset.Tables | Select-Object -Expand Rows)
    
        Return $rows
    
    }

}

