import { mapAttributes, mapTables } from "../src/lib/metadata";
import type { AttributeInfo, TableInfo } from "../src/lib/types";

const label = (text: string) => ({ UserLocalizedLabel: { Label: text }, LocalizedLabels: [{ Label: text, LanguageCode: 1033 }] });

const attr = (logical: string, display: string, typeName: string, extra: Record<string, unknown> = {}) => ({
    LogicalName: logical,
    SchemaName: logical.replace(/(^|_)([a-z])/g, (_m, p, c) => p + c.toUpperCase()),
    DisplayName: label(display),
    AttributeType: typeName.replace(/Type$/, ""),
    AttributeTypeName: { Value: typeName },
    AttributeOf: null,
    IsValidForRead: true,
    IsPrimaryId: false,
    IsPrimaryName: false,
    IsCustomAttribute: logical.startsWith("new_"),
    ...extra,
});

export const rawTables = [
    {
        LogicalName: "account",
        DisplayName: label("Account"),
        EntitySetName: "accounts",
        PrimaryIdAttribute: "accountid",
        PrimaryNameAttribute: "name",
        IsCustomEntity: false,
        IsIntersect: false,
        IsActivity: false,
        IsPrivate: false,
    },
    {
        LogicalName: "contact",
        DisplayName: label("Contact"),
        EntitySetName: "contacts",
        PrimaryIdAttribute: "contactid",
        PrimaryNameAttribute: "fullname",
        IsCustomEntity: false,
        IsIntersect: false,
        IsActivity: false,
        IsPrivate: false,
    },
    {
        LogicalName: "new_project",
        DisplayName: label("Project"),
        EntitySetName: "new_projects",
        PrimaryIdAttribute: "new_projectid",
        PrimaryNameAttribute: "new_name",
        IsCustomEntity: true,
        IsIntersect: false,
        IsActivity: false,
        IsPrivate: false,
    },
    { LogicalName: "privatething", DisplayName: label("Private"), EntitySetName: "privatethings", PrimaryIdAttribute: "privatethingid", IsPrivate: true },
    { LogicalName: "noset", DisplayName: label("No set"), EntitySetName: null, PrimaryIdAttribute: "nosetid" },
];

export const rawAccountAttributes = [
    attr("accountid", "Account", "UniqueidentifierType", { IsPrimaryId: true }),
    attr("name", "Account Name", "StringType", { IsPrimaryName: true }),
    attr("primarycontactid", "Primary Contact", "LookupType"),
    attr("primarycontactidname", "Primary Contact Name", "StringType", { AttributeOf: "primarycontactid" }),
    attr("ownerid", "Owner", "OwnerType"),
    attr("industrycode", "Industry", "PicklistType"),
    attr("statecode", "Status", "StateType"),
    attr("statuscode", "Status Reason", "StatusType"),
    attr("revenue", "Annual Revenue", "MoneyType"),
    attr("numberofemployees", "Number of Employees", "IntegerType"),
    attr("donotemail", "Do not allow Emails", "BooleanType"),
    attr("createdon", "Created On", "DateTimeType"),
    attr("new_startdate", "Start Date", "DateTimeType"),
    attr("new_reviewdate", "Review Date", "DateTimeType"),
    attr("description", "Description", "MemoType"),
    attr("entityimage", "Default Image", "ImageType"),
    attr("new_tags", "Tags", "MultiSelectPicklistType"),
    attr("hidden", "Hidden", "StringType", { IsValidForRead: false }),
];

export const rawAccountDates = [
    { LogicalName: "createdon", Format: "DateAndTime", DateTimeBehavior: { Value: "UserLocal" } },
    { LogicalName: "new_startdate", Format: "DateOnly", DateTimeBehavior: { Value: "DateOnly" } },
    { LogicalName: "new_reviewdate", Format: "DateAndTime", DateTimeBehavior: { Value: "TimeZoneIndependent" } },
];

export const rawAccountLookups = [
    { LogicalName: "primarycontactid", Targets: ["contact"] },
    { LogicalName: "ownerid", Targets: ["systemuser", "team"] },
];

export const rawContactAttributes = [
    attr("contactid", "Contact", "UniqueidentifierType", { IsPrimaryId: true }),
    attr("fullname", "Full Name", "StringType", { IsPrimaryName: true }),
    attr("parentcustomerid", "Company Name", "CustomerType"),
    attr("emailaddress1", "Email", "StringType"),
    attr("gendercode", "Gender", "PicklistType"),
];

export function fixture(): { tables: Map<string, TableInfo>; attributes: Map<string, AttributeInfo[]>; tableList: TableInfo[] } {
    const tableList = mapTables(rawTables);
    const tables = new Map(tableList.map((t) => [t.logicalName, t]));
    const attributes = new Map<string, AttributeInfo[]>([
        ["account", mapAttributes(rawAccountAttributes, { dateTimes: rawAccountDates, lookups: rawAccountLookups })],
        ["contact", mapAttributes(rawContactAttributes)],
    ]);
    return { tables, attributes, tableList };
}
