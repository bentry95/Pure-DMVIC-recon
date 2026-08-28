-- =========================================================
-- Sample MS SQL Server Script
-- =========================================================

-- Batch 1: Print starting message
PRINT N'Starting SQL Script Execution Test...';
GO

-- Batch 2: Create temporary table and insert sample rows
IF OBJECT_ID('tempdb..#SampleUsers') IS NOT NULL
    DROP TABLE #SampleUsers;
GO

CREATE TABLE #SampleUsers (
    ID INT IDENTITY(1,1) PRIMARY KEY,
    Username NVARCHAR(50) NOT NULL,
    Email NVARCHAR(100) NOT NULL,
    CreatedDate DATETIME DEFAULT GETDATE()
);
GO

-- Batch 3: Insert sample data
INSERT INTO #SampleUsers (Username, Email)
VALUES 
    (N'john_doe', N'john@example.com'),
    (N'jane_smith', N'jane@example.com'),
    (N'admin_user', N'admin@example.com');
GO

-- Batch 4: Query inserted records
SELECT * FROM #SampleUsers;
GO

-- Batch 5: Cleanup temporary table
DROP TABLE #SampleUsers;
PRINT N'SQL Script Execution Test Completed Successfully!';
GO
