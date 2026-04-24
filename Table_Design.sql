-- 1) Project master
CREATE TABLE core.Project_Master (
    ProjectID           NVARCHAR(50)   NOT NULL,
    ProjectClass        NVARCHAR(20)   NULL,
    TG1_Date            DATE           NULL,
    TG3_Estimated_Date  DATE           NULL,
    TG3_Actual_Date     DATE           NULL,
    TG4_Estimated_Date  DATE           NULL,
    TG4_Actual_Date     DATE           NULL,
    UpdatedAt           DATETIME2(0)   NOT NULL,
    CONSTRAINT PK_Project_Master PRIMARY KEY (ProjectID)
);
GO

-- 2) Labor hours summary
CREATE TABLE core.Fact_LaborHours (
    ProjectID         NVARCHAR(50)    NOT NULL,
    Estimated_Hours   DECIMAL(18,2)   NOT NULL DEFAULT 0,
    Actual_Hours      DECIMAL(18,2)   NOT NULL DEFAULT 0,
    Variance_Hours    DECIMAL(18,2)   NULL,
    Variance_Percent  DECIMAL(18,6)   NULL,
    UpdatedAt         DATETIME2(0)    NOT NULL,
    CONSTRAINT PK_Fact_LaborHours PRIMARY KEY (ProjectID)
);
GO

-- 3) Team hours combined table
CREATE TABLE core.Fact_TeamHours (
    ProjectID         NVARCHAR(50)    NOT NULL,
    FunctionGroup     NVARCHAR(50)    NOT NULL,
    Estimated_Hours   DECIMAL(18,2)   NOT NULL DEFAULT 0,
    Actual_Hours      DECIMAL(18,2)   NOT NULL DEFAULT 0,
    UpdatedAt         DATETIME2(0)    NOT NULL,
    CONSTRAINT PK_Fact_TeamHours PRIMARY KEY (ProjectID, FunctionGroup)
);
GO

-- 4) Cost table
CREATE TABLE core.Fact_Costs (
    ProjectID                NVARCHAR(50)    NOT NULL,
    TG1_Material_Cost        DECIMAL(18,2)   NULL,
    TG1_DevCost_NoNRE        DECIMAL(18,2)   NULL,
    TG1_NPI_Cost             DECIMAL(18,2)   NULL,
    TG1_Forecasted_Budget    DECIMAL(18,2)   NULL,
    TG3_Material_Cost        DECIMAL(18,2)   NULL,
    TG3_DevCost_NoNRE        DECIMAL(18,2)   NULL,
    TG3_NPI_Cost             DECIMAL(18,2)   NULL,
    TG3_Actual_Cost          DECIMAL(18,2)   NULL,
    TG4_Material_Cost        DECIMAL(18,2)   NULL,
    TG4_DevCost_NoNRE        DECIMAL(18,2)   NULL,
    TG4_NPI_Cost             DECIMAL(18,2)   NULL,
    TG4_Actual_Cost          DECIMAL(18,2)   NULL,
    UpdatedAt                DATETIME2(0)    NOT NULL,
    CONSTRAINT PK_Fact_Costs PRIMARY KEY (ProjectID)
);
GO

-- 5) Scope clusters
CREATE TABLE core.Fact_ScopeClusters (
    ProjectID             NVARCHAR(50)    NOT NULL,
    ScopeCluster          NVARCHAR(100)   NULL,
    KeyFunctionsInvolved  NVARCHAR(255)   NULL,
    ClusterDescription    NVARCHAR(MAX)   NULL,
    UpdatedAt             DATETIME2(0)    NOT NULL,
    CONSTRAINT PK_Fact_ScopeClusters PRIMARY KEY (ProjectID)
);
GO

-- 6) Dimension tables
CREATE TABLE core.DIM_ProjectID (
    ProjectID   NVARCHAR(50)   NOT NULL,
    CONSTRAINT PK_DIM_ProjectID PRIMARY KEY (ProjectID)
);
GO

CREATE TABLE core.DIM_ProjectClass (
    ProjectClass   NVARCHAR(20)   NOT NULL,
    CONSTRAINT PK_DIM_ProjectClass PRIMARY KEY (ProjectClass)
);
GO

CREATE TABLE core.DIM_FunctionGroup (
    FunctionGroup   NVARCHAR(50)   NOT NULL,
    CONSTRAINT PK_DIM_FunctionGroup PRIMARY KEY (FunctionGroup)
);
GO
