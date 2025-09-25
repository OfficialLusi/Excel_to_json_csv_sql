# ddl_module.py

from pathlib import Path  # FS paths
from typing import Dict, Any, List, Optional  # typing
import json  # JSON read
import re  # regex

MaxIdentLen = 30  # Oracle identifier length limit 

def SanitizeIdentifier(name: str) -> str:
    s = re.sub(r"\s+", "_", name.strip())  # spaces → underscore 
    s = re.sub(r"[^A-Za-z0-9_]", "", s).upper()  # keep only A-Z0-9_ and uppercase 
    if not s:
        s = "C_COL"  # fallback name 
    if s[0].isdigit():
        s = "C_" + s  # cannot start with digit in Oracle without quotes 
    return s[:MaxIdentLen]  # enforce length 

def MakeConstraintName(prefix: str, table: str, suffix: str = "") -> str:
    base = f"{prefix}_{table}" if not suffix else f"{prefix}_{table}_{suffix}"  # compose base 
    return SanitizeIdentifier(base)[:MaxIdentLen]  # sanitize and trim 

def IsTrue(value: Any) -> bool:
    if value is None:
        return False  # None is false 
    s = str(value).strip().casefold()  # normalized string 
    return s in {"y", "yes", "true", "1", "pk", "si", "sì", "x"}  # accepted truthy tokens 

def OracleQuoteLiteral(value: Any) -> str:
    s = str(value).replace("'", "''")  # escape single quotes 
    return f"'{s}'"  # quote as SQL string 

def NormalizeDefault(defaultValue: Any, dataType: str) -> Optional[str]:
    if defaultValue is None:
        return None  # no default 
    s = str(defaultValue).strip()  # trim 
    if s == "" or s == "-" or s.casefold() == "null":
        return None  # treat as no default 
    try:
        float(s)  # numeric literal 
        return s  # keep numeric (no quotes) 
    except Exception:
        pass
    dtUpper = (dataType or "").upper()  # type upper 
    if "TIMESTAMP" in dtUpper or dtUpper == "DATE":
        return OracleQuoteLiteral(s)  # keep quoted (optionally TO_DATE/TO_TIMESTAMP) 
    if s.casefold() in {"true", "false"}:
        return "1" if s.casefold() == "true" else "0"  # boolean as 1/0 
    return OracleQuoteLiteral(s)  # fallback quoted 

def GenerateTableDdl(tableObj: Dict[str, Any], schema: str) -> str:
    RawTableName = str(tableObj["table_name"])  # raw name 
    TableName = SanitizeIdentifier(RawTableName)  # sanitized table name 
    Rows = tableObj["rows"]  # data rows 
    ColumnDefs: List[str] = []  # column definitions 
    PkCols: List[str] = []  # PK columns 
    ColumnComments: List[str] = []  # comments on columns 
    def Get(row: Dict[str, Any], key: str, default=None):
        return row.get(key, default)  # safe get 
    for r in Rows:  # each Excel-defined column 
        RawColName = Get(r, "COLUMN NAME")  # column name 
        if not RawColName:
            continue  # skip invalid row 
        ColName = SanitizeIdentifier(str(RawColName))  # sanitized column name 
        DataType = str(Get(r, "DATA TYPE", "VARCHAR2(4000)")).strip()  # Oracle datatype 
        NullFlag = Get(r, "NULL")  # NULL flag 
        NotNull = str(NullFlag).strip().casefold() in {"no", "n", "false", "0"}  # evaluate not-null 
        DefaultExpr = NormalizeDefault(Get(r, "DEFAULT"), DataType)  # default expression 
        DefaultClause = f" DEFAULT {DefaultExpr}" if DefaultExpr is not None else ""  # default clause 
        NullClause = " NOT NULL" if NotNull else ""  # nullability clause 
        ColumnDefs.append(f"  {ColName} {DataType}{DefaultClause}{NullClause}")  # append column line 
        if IsTrue(Get(r, "PK")): # if primary key
            PkCols.append(ColName)  # collect PK column 
        Description = Get(r, "DESCRIPTION")  # column description 
        if Description and str(Description).strip():
            Txt = str(Description).replace("'", "''")  # escape quotes 
            ColumnComments.append(f"COMMENT ON COLUMN {schema}.{TableName}.{ColName} IS '{Txt}';")  # comment line 
    DdlLines: List[str] = []  # final DDL 
    DdlLines.append(f"CREATE TABLE {schema}.{TableName} (")  # create table start 
    DdlLines.append(",\n".join(ColumnDefs))  # columns 
    DdlLines.append(");")  # create table end 
    if PkCols:  # add PK if present 
        PkName = MakeConstraintName("PK", TableName)  # PK name 
        ColsCsv = ", ".join(PkCols)  # pk columns 
        DdlLines.append(f"ALTER TABLE {schema}.{TableName} ADD CONSTRAINT {PkName} PRIMARY KEY ({ColsCsv});")  # PK ddl 
    # optional table comment inferred from uniform COMMENTS 
    TableComment = tableObj.get("comments")  # explicit comment if provided 
    if not TableComment:
        CommentsVals = {str(r.get("COMMENTS")).strip() for r in Rows if r.get("COMMENTS")}  # unique comments 
        if len(CommentsVals) == 1:
            TableComment = CommentsVals.pop()  # use if unique 
    if TableComment:
        Txt = str(TableComment).replace("'", "''")  # escape quotes 
        DdlLines.append(f"COMMENT ON TABLE {schema}.{TableName} IS '{Txt}';")  # table comment 
    DdlLines.extend(ColumnComments)  # append column comments 
    return "\n".join(DdlLines) + "\n"  # final DDL string 

def WriteAllDdls(tables: List[Dict[str, Any]], schemaBySheet: Dict[str, str], outDir: Path):
    outDir.mkdir(parents=True, exist_ok=True)  # ensure folder 
    allDdls: List[str] = []  # aggregate ddl 
    for t in tables:  # per-table 
        sheetName = t.get("sheet", "")  # source sheet 
        schema = schemaBySheet.get(sheetName, next(iter(schemaBySheet.values())))  # schema by sheet or default 
        rawName = str(t["table_name"])  # raw table name #
        fileStem = SanitizeIdentifier(rawName) or "TABLE"  # filename stem 
        sqlPath = outDir / f"{fileStem}.sql"  # file path 
        ddl = GenerateTableDdl(t, schema)  # generate DDL 
        sqlPath.write_text(ddl, encoding="utf-8")  # write 
        allDdls.append(ddl)  # collect 
    (outDir / "_ALL_TABLES.sql").write_text("\n".join(allDdls), encoding="utf-8")  # aggregate file 

def LoadTablesFromJson(jsonPath: Path) -> List[Dict[str, Any]]:
    return json.loads(jsonPath.read_text(encoding="utf-8"))  # load tables list 
