#region  == Copyright == 
// =====================================================================
//  Microsoft Consulting Services France - Aymeric Mouillé - 2013
//  Projet     : CRM.Model.Generator - CRM.Model.Generator
//  Fichier    : Program.cs (01/05/2013 13:18)
//  
//  Copyright (C) Microsoft Corporation.  All rights reserved.
// =====================================================================
#endregion

namespace Crm.Model.Generator
{
    #region  == Using == 
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using Microsoft.Xrm.Client;
    using Microsoft.Xrm.Client.Services;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;

    #endregion

    internal class Program
    {
        private static string AttributeCodeTemplate;

        private static string DefaultNamespace;

        private static string EntityBaseType;

        private static string EntityClassCodeTemplate;

        private static string GlobalOptionSetClassCodeTemplate;

        private static string GlobalOptionSetCodeTemplate;

        private static string AttributePropertyTemplate;

        private static string OptionSetEnumCodeTemplate;

        private static string TargetPath;

        private static OrganizationService organizationService;

        #region Main
        /// <summary>
        /// Program entry point
        /// </summary>
        /// <param name="args"></param>
        private static void Main(string[] args)
        {
            try
            {
                string outputPath = ConfigurationManager.AppSettings["Folder.Path"];
                TargetPath = Path.Combine(outputPath, "Crm.Model.Data");
                DefaultNamespace = ConfigurationManager.AppSettings["DefaultNamespace"];
                EntityBaseType = ConfigurationManager.AppSettings["EntityBaseType"];

                if (Directory.Exists(TargetPath))
                {
                    try
                    {
                        Directory.Delete(TargetPath, true);
                    }
                    catch (Exception)
                    {
                        Thread.Sleep(2000);
                    }

                    if (Directory.Exists(TargetPath))
                    {
                        try
                        {
                            Directory.Delete(TargetPath, true);
                        }
                        catch (Exception)
                        {
                            Thread.Sleep(2000);
                        }
                    }
                }
                Thread.Sleep(2000);

                if (!Directory.Exists(TargetPath))
                {
                    Directory.CreateDirectory(TargetPath);
                }
                Thread.Sleep(2000);

                Directory.CreateDirectory(Path.Combine(TargetPath, "Custom"));
                Directory.CreateDirectory(Path.Combine(TargetPath, "System"));

                EntityClassCodeTemplate = LoadTemplateCode(TemplateNames.EntityClassCodeTemplate);
                AttributeCodeTemplate = LoadTemplateCode(TemplateNames.AttributeCodeTemplate);
                GlobalOptionSetClassCodeTemplate = LoadTemplateCode(TemplateNames.GlobalOptionSetClassCodeTemplate);
                GlobalOptionSetCodeTemplate = LoadTemplateCode(TemplateNames.GlobalOptionSetCodeTemplate);
                OptionSetEnumCodeTemplate = LoadTemplateCode(TemplateNames.OptionSetEnumCodeTemplate);
                AttributePropertyTemplate = LoadTemplateCode(TemplateNames.AttributePropertyTemplate);

                Console.WriteLine("Connection : initializing ...");
                InitializeConnection();
                Console.WriteLine("Connection : initialized!");

                Console.WriteLine("Option Set : Loading ...");
                LoadGlobalOptionSet();
                Console.WriteLine("Option Set : Loaded!");

                Console.WriteLine("Metadata : Loading ...");
                var entities = RetrieveEntityMetadata();
                Console.WriteLine("Metadata : Loaded!");
                Console.WriteLine("Entities : processing ...");

                StringBuilder csvExtract = new StringBuilder();
                csvExtract.AppendLine("Entité, Table, Code, Nom complet, Nom, Type, N, R, S, A, Règle, Taille");
                foreach (EntityMetadata currentEntity in entities)
                {
                    Console.WriteLine("  Entity : {0}", currentEntity.LogicalName);

                    string entityDisplayName = currentEntity.SchemaName;
                    if (currentEntity.DisplayName.UserLocalizedLabel != null)
                    {
                        entityDisplayName = currentEntity.DisplayName.UserLocalizedLabel.Label;
                    }

                    foreach (var attribute in currentEntity.Attributes)
                    {
                        string attributeDisplayName = attribute.SchemaName;
                        if (attribute.DisplayName.UserLocalizedLabel != null)
                        {
                            attributeDisplayName = attribute.DisplayName.UserLocalizedLabel.Label;
                        }

                        csvExtract.AppendLine(string.Format("{0}, {1}, N/A, {2}, {3}, {4}, {5}, {6}, {7}, {8}, N/A, {9}", entityDisplayName, // 0 : Entité
                            currentEntity.LogicalName, // 1 : Table
                            attributeDisplayName, // 2 : Nom complet
                            attribute.LogicalName, // 3 : Nom
                            attribute.AttributeType.Value, // 4 : Type
                            (attribute.RequiredLevel.Value == AttributeRequiredLevel.SystemRequired) ? "O" : "N", // 5 : N (Niveau requis)
                            attribute.IsValidForAdvancedFind.Value ? "O" : "N", // 6 : R (Recherche)
                            attribute.IsSecured.Value ? "O" : "N", // 7 : S (Sécurité de niveau champ) 
                            attribute.IsAuditEnabled.Value ? "O" : "N", // 8 : A (Audit)
                            GetAttributeLength(attribute) // 9
                            ));
                    }

                    string entityCustomClassName = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(currentEntity.SchemaName);

                    string entityCustomName = string.Concat("Crm", Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(currentEntity.SchemaName));
                    string path = TargetPath;
                    if (currentEntity.IsCustomEntity.HasValue
                        && currentEntity.IsCustomEntity.Value)
                    {
                        path += @"\Custom\";
                    }
                    else
                    {
                        path += @"\System\";
                    }
                    string fileName = Path.Combine(path, string.Format("{0}.generated.cs", entityCustomName));
                    File.WriteAllText(fileName, GenerateEntityClassCode(currentEntity, entityCustomName, EntityBaseType));
                }

                string csvFileName = Path.Combine(TargetPath, "MetaData.csv");
                File.WriteAllText(csvFileName, csvExtract.ToString(), Encoding.UTF8);

                Console.WriteLine("Entities : processed!");
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("CRM.Model.Generator", ex.ToString(), EventLogEntryType.Error);
            }
            finally
            {
                organizationService.Dispose();
                Console.WriteLine("Done.");
            }
        }
        #endregion

        #region Retrieve Entity Metadata
        /// <summary>
        /// Get all published entities metadata from server
        /// </summary>
        /// <returns></returns>
        private static EntityMetadata[] RetrieveEntityMetadata()
        {
            RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest { EntityFilters = EntityFilters.Attributes, RetrieveAsIfPublished = true };
            RetrieveAllEntitiesResponse response = (RetrieveAllEntitiesResponse)organizationService.Execute(request);
            return response.EntityMetadata;
        }
        #endregion

        #region Generate Entity Class Code
        /// <summary>
        /// Output entity metadata information to class definition content
        /// </summary>
        /// <param name="currentEntity">The entity metadata object</param>
        /// <param name="entityCustomName">The entity name for definition class</param>
        /// <returns>Entity definition class code</returns>
        private static string GenerateEntityClassCode(EntityMetadata currentEntity, string entityCustomName, string baseType)
        {
            string entityClassCode = EntityClassCodeTemplate;
            entityClassCode = entityClassCode.Replace("[@DefaultNamespace]", DefaultNamespace);
            entityClassCode = entityClassCode.Replace("[@EntityCustomName]", entityCustomName);
            entityClassCode = entityClassCode.Replace("[@Entity.SchemaName]", currentEntity.SchemaName);
            entityClassCode = entityClassCode.Replace("[@Entity.LogicalName]", currentEntity.LogicalName);

            baseType = String.IsNullOrWhiteSpace(baseType) ? "" : $"{baseType}, ";

            entityClassCode = entityClassCode.Replace("[@EntityBaseType]", baseType);

            string description = currentEntity.SchemaName;
            if (currentEntity.Description.UserLocalizedLabel != null)
            {
                description = currentEntity.Description.UserLocalizedLabel.Label;
            }
            entityClassCode = entityClassCode.Replace("[@EntityDescription]", description);

            if (currentEntity.ObjectTypeCode.HasValue)
            {
                entityClassCode = entityClassCode.Replace("[@Entity.ObjectTypeCode.Value]", currentEntity.ObjectTypeCode.Value.ToString());
            }

            var customAttributes = currentEntity.Attributes.Where(a => a.IsCustomAttribute == true).OrderBy(a => a.SchemaName).ToList();
            customAttributes = customAttributes.Where(a => a.AttributeOf == null).ToList();
            customAttributes = customAttributes.Where(a => a.AttributeType != null).ToList();
            customAttributes = customAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Customer).ToList();
            customAttributes = customAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Owner).ToList();
            customAttributes = customAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Lookup).ToList();
            customAttributes = customAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Picklist).ToList();
            customAttributes = customAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Virtual).ToList();
            customAttributes = customAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.BigInt).ToList();
            customAttributes = customAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.ManagedProperty).ToList();
            customAttributes = customAttributes.Where(a => !a.LogicalName.EndsWith("_base")).ToList();
            string customAttributesCode = LoadAttributes(customAttributes);
            entityClassCode = entityClassCode.Replace("[@CustomAttributes]", customAttributesCode);

            string customPropertiesCode = LoadProperties(customAttributes);
            entityClassCode = entityClassCode.Replace("[@CustomAttributeProperties]", customPropertiesCode);

            var systemAttributes = currentEntity.Attributes.Where(a => a.IsCustomAttribute == false).OrderBy(a => a.SchemaName).ToList();
            systemAttributes = systemAttributes.Where(a => a.AttributeOf == null).ToList();
            systemAttributes = systemAttributes.Where(a => a.AttributeType != null).ToList();
            systemAttributes = systemAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Customer).ToList();
            systemAttributes = systemAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Owner).ToList();
            systemAttributes = systemAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Lookup).ToList();
            systemAttributes = systemAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Picklist).ToList();
            systemAttributes = systemAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Virtual).ToList();
            systemAttributes = systemAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.BigInt).ToList();
            systemAttributes = systemAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.ManagedProperty).ToList();
            systemAttributes = systemAttributes.Where(a => a.AttributeType.Value != AttributeTypeCode.Customer).ToList();
            systemAttributes = systemAttributes.Where(a => !a.LogicalName.EndsWith("_base")).ToList();
            string systemAttributesCode = LoadAttributes(systemAttributes);
            entityClassCode = entityClassCode.Replace("[@SystemAttributes]", systemAttributesCode);

            string systemPropertiesCode = LoadProperties(systemAttributes);
            entityClassCode = entityClassCode.Replace("[@SystemAttributeProperties]", systemPropertiesCode);

            var lookupsAttributes = currentEntity.Attributes.OrderBy(a => a.SchemaName).ToList();
            lookupsAttributes = lookupsAttributes.Where(a => a.AttributeType != null).ToList();
            lookupsAttributes = lookupsAttributes.Where(a => a.AttributeType.Value == AttributeTypeCode.Lookup || a.AttributeType.Value == AttributeTypeCode.Customer || a.AttributeType.Value == AttributeTypeCode.Owner).ToList();
            string lookupsAttributesCode = LoadAttributes(lookupsAttributes);
            entityClassCode = entityClassCode.Replace("[@Lookups]", lookupsAttributesCode);

            var optionSetAttributes = currentEntity.Attributes.OrderBy(a => a.SchemaName).ToList();
            optionSetAttributes = optionSetAttributes.Where(a => a.AttributeType != null).ToList();
            optionSetAttributes = optionSetAttributes.Where(a => a.AttributeType.Value == AttributeTypeCode.Picklist).ToList();
            optionSetAttributes = optionSetAttributes.Where(a => a.AttributeOf == null).ToList();
            string optionSetAttributesCode = LoadAttributes(optionSetAttributes);
            entityClassCode = entityClassCode.Replace("[@OptionSets]", optionSetAttributesCode);

            return entityClassCode;
        }

        #endregion

        #region Get Attribute Length
        /// <summary>
        /// Get Attribute Length
        /// </summary>
        /// <param name="attribute"></param>
        /// <returns></returns>
        private static string GetAttributeLength(AttributeMetadata attribute)
        {
            if (attribute.AttributeType.Value == AttributeTypeCode.String)
            {
                var convertedAttribute = (StringAttributeMetadata)attribute;
                if (convertedAttribute.MaxLength.HasValue)
                {
                    return convertedAttribute.MaxLength.Value.ToString();
                }
            }

            if (attribute.AttributeType.Value == AttributeTypeCode.BigInt)
            {
                var convertedAttribute = (BigIntAttributeMetadata)attribute;
                if (convertedAttribute.MaxValue.HasValue)
                {
                    return convertedAttribute.MaxValue.Value.ToString();
                }
            }

            if (attribute.AttributeType.Value == AttributeTypeCode.Decimal)
            {
                var convertedAttribute = (DecimalAttributeMetadata)attribute;
                if (convertedAttribute.MaxValue.HasValue)
                {
                    string value = convertedAttribute.MaxValue.Value.ToString();
                    if (convertedAttribute.Precision.HasValue)
                    {
                        value += string.Format(" (Precision : {0})", convertedAttribute.Precision.Value);
                    }
                    return value;
                }
            }

            if (attribute.AttributeType.Value == AttributeTypeCode.Double)
            {
                var convertedAttribute = (DoubleAttributeMetadata)attribute;
                if (convertedAttribute.MaxValue.HasValue)
                {
                    string value = convertedAttribute.MaxValue.Value.ToString();
                    if (convertedAttribute.Precision.HasValue)
                    {
                        value += string.Format(" (Precision : {0})", convertedAttribute.Precision.Value);
                    }
                    return value;
                }
            }

            if (attribute.AttributeType.Value == AttributeTypeCode.Money)
            {
                var convertedAttribute = (MoneyAttributeMetadata)attribute;
                if (convertedAttribute.MaxValue.HasValue)
                {
                    string value = convertedAttribute.MaxValue.Value.ToString();
                    if (convertedAttribute.Precision.HasValue)
                    {
                        value += string.Format(" (Precision : {0})", convertedAttribute.Precision.Value);
                    }
                    return value;
                }
            }

            if (attribute.AttributeType.Value == AttributeTypeCode.Memo)
            {
                var convertedAttribute = (MemoAttributeMetadata)attribute;
                if (convertedAttribute.MaxLength.HasValue)
                {
                    return ((MemoAttributeMetadata)attribute).MaxLength.Value.ToString();
                }
            }

            if (attribute.AttributeType.Value == AttributeTypeCode.Integer)
            {
                var convertedAttribute = (IntegerAttributeMetadata)attribute;
                if (convertedAttribute.MaxValue.HasValue)
                {
                    return ((IntegerAttributeMetadata)attribute).MaxValue.Value.ToString();
                }
            }

            return string.Empty;
        }
        #endregion

        #region Initialize Connection
        /// <summary>
        /// Initialize proxy with server connection settings
        /// </summary>
        private static void InitializeConnection()
        {
            string crmConnectionString = ConfigurationManager.AppSettings["Crm.ConnectionString"];
            CrmConnection crmConnection = CrmConnection.Parse(crmConnectionString);
            organizationService = new OrganizationService(crmConnection);
        }
        #endregion

        #region Load Attributes
        /// <summary>
        /// Generate constant strings with given attribute collection
        /// </summary>
        /// <param name="attributes">Attribute list</param>
        /// <returns>Generated code</returns>
        private static string LoadAttributes(List<AttributeMetadata> attributes)
        {
            string attributesCode = string.Empty;
            foreach (AttributeMetadata currentAttribute in attributes)
            {
                string attributeCode = string.Empty;

                string displayName = currentAttribute.SchemaName;
                if (currentAttribute.DisplayName.UserLocalizedLabel != null)
                {
                    displayName = currentAttribute.DisplayName.UserLocalizedLabel.Label;
                }

                string description = displayName;
                if (currentAttribute.Description.UserLocalizedLabel != null)
                {
                    description = currentAttribute.Description.UserLocalizedLabel.Label;
                }

                if (currentAttribute.AttributeType.Value == AttributeTypeCode.Picklist)
                {
                    attributeCode = GlobalOptionSetCodeTemplate;
                }
                else
                {
                    attributeCode = AttributeCodeTemplate;
                }
                attributeCode = attributeCode.Replace("[@Attribute.Description]", TransformToSymmary(description));
                attributeCode = attributeCode.Replace("[@Attribute.DisplayName]", displayName);
                attributeCode = attributeCode.Replace("[@Attribute.SchemaName]", currentAttribute.SchemaName);
                attributeCode = attributeCode.Replace("[@Attribute.LogicalName]", currentAttribute.LogicalName);
                attributeCode = attributeCode.Replace("[@Attribute.AttributeType.Value]", currentAttribute.AttributeType.Value.ToString());

                string validity = string.Empty;
                validity += currentAttribute.IsValidForRead.Value ? " Read |" : String.Empty;
                validity += currentAttribute.IsValidForCreate.Value ? " Create |" : String.Empty;
                validity += currentAttribute.IsValidForUpdate.Value ? " Update |" : String.Empty;
                validity += currentAttribute.IsValidForAdvancedFind.Value ? " AdvancedFind |" : String.Empty;
                if (validity.EndsWith(" |"))
                {
                    validity = validity.Remove(validity.LastIndexOf('|'));
                }
                attributeCode = attributeCode.Replace("[@Attribute.Validity]", validity);

                if (currentAttribute.AttributeType.Value == AttributeTypeCode.Picklist)
                {
                    attributeCode = attributeCode.Replace("[@OptionSet.Description]", TransformToSymmary(description));
                    attributeCode = attributeCode.Replace("[@OptionSet.DisplayName]", displayName);
                    attributeCode = attributeCode.Replace("[@OptionSet.SchemaName]", currentAttribute.SchemaName);
                    attributeCode = attributeCode.Replace("[@OptionSet.LogicalName]", currentAttribute.LogicalName);
                    attributeCode = attributeCode.Replace("[@OptionSet.OptionSetType.Value]", currentAttribute.AttributeType.Value.ToString());

                    string optionSetEnums = string.Empty;
                    PicklistAttributeMetadata optionSetMetadata = (PicklistAttributeMetadata)currentAttribute;
                    int optionCount = 1;
                    foreach (OptionMetadata option in optionSetMetadata.OptionSet.Options)
                    {
                        var desc = option.Label.UserLocalizedLabel.Label;
                        var label = optionCount + "_" + ConvertNameAsVariable(desc);
                        var value = option.Value.Value.ToString(CultureInfo.InvariantCulture);

                        string optionSetEnumCode = OptionSetEnumCodeTemplate;
                        optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Description]", TransformToSymmary(desc));
                        optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Label]", label);
                        optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Value]", value);
                        optionSetEnums += optionSetEnumCode;
                        optionCount++;
                    }
                    attributeCode = attributeCode.Replace("[@OptionSet.Values]", optionSetEnums);
                }

                attributesCode += attributeCode;
            }
            return attributesCode;
        }


        private static string LoadProperties(List<AttributeMetadata> attributes)
        {
            List<String> propertyCodes = new List<string>();

            foreach (AttributeMetadata attr in attributes)
            {
                var proppertyCode = AttributePropertyTemplate;


                string displayName = attr.SchemaName;
                if (attr.DisplayName.UserLocalizedLabel != null)
                {
                    displayName = attr.DisplayName.UserLocalizedLabel.Label;
                }

                string description = displayName;
                if (attr.Description.UserLocalizedLabel != null)
                {
                    description = attr.Description.UserLocalizedLabel.Label;
                }
                
                proppertyCode = proppertyCode.Replace("[@Attribute.Description]", TransformToSymmary(description));
                proppertyCode = proppertyCode.Replace("[@Attribute.LogicalName]", attr.LogicalName);
                proppertyCode = proppertyCode.Replace("[@Attribute.Type]", GetCsType(attr));
                proppertyCode = proppertyCode.Replace("[@Attribute.SchemaName]", attr.SchemaName);


                propertyCodes.Add(proppertyCode);
            }

            return String.Join("\r\n", propertyCodes);
        }

        private static string GetCsType(AttributeMetadata attr)
        {
            switch (attr.AttributeType.Value)
            {
                case AttributeTypeCode.Boolean:
                    return "bool?";
                case AttributeTypeCode.DateTime:
                    return "DateTime?";
                case AttributeTypeCode.Decimal:
                case AttributeTypeCode.Money:
                    return "decimal?";
                case AttributeTypeCode.Double:
                    return "double?";
                case AttributeTypeCode.Integer:
                    return "int?";
                case AttributeTypeCode.Lookup:
                    return "Microsoft.Xrm.Sdk.EntityReference";
                case AttributeTypeCode.Memo:
                case AttributeTypeCode.String:
                case AttributeTypeCode.EntityName:
                    return "string";
                case AttributeTypeCode.Picklist:
                case AttributeTypeCode.State:
                case AttributeTypeCode.Status:
                    return "Microsoft.Xrm.Sdk.OptionSetValue";
                case AttributeTypeCode.Uniqueidentifier:
                    return "Guid?";
                case AttributeTypeCode.PartyList:
                    return "Microsoft.Xrm.Sdk.EntityCollection";

                default:
                    return attr.AttributeType.Value.ToString();
                //throw new NotSupportedException($"AttributeType {attr.AttributeType.Value} not supported");
            }
        }

        #endregion

        #region Load Global Option Set
        /// <summary>
        /// Load Global Option Set
        /// </summary>
        private static void LoadGlobalOptionSet()
        {
            RetrieveAllOptionSetsRequest request = new RetrieveAllOptionSetsRequest { };
            RetrieveAllOptionSetsResponse response = (RetrieveAllOptionSetsResponse)organizationService.Execute(request);
            var optionSetData = response.OptionSetMetadata;
            
            string content = string.Empty;
            foreach (var optionSet in optionSetData)
            {
                if (optionSet.IsGlobal.HasValue
                    && !optionSet.IsGlobal.Value)
                {
                    continue;
                }

                string optionSetCode = GlobalOptionSetCodeTemplate;

                string label = GetOptionSetLabel(optionSet);
                string displayName = label;
                string description = label;
                if (string.IsNullOrEmpty(description))
                {
                    description = displayName;
                }

                optionSetCode = optionSetCode.Replace("[@OptionSet.DisplayName]", displayName);
                optionSetCode = optionSetCode.Replace("[@OptionSet.Description]", TransformToSymmary(description));
                optionSetCode = optionSetCode.Replace("[@OptionSet.SchemaName]", optionSet.Name);
                optionSetCode = optionSetCode.Replace("[@OptionSet.OptionSetType.Value]", optionSet.OptionSetType.Value.ToString());
                optionSetCode = optionSetCode.Replace("[@OptionSet.LogicalName]", optionSet.Name);

                string optionSetEnums = string.Empty;

                if (optionSet.OptionSetType != null)
                {
                    if ((OptionSetType)optionSet.OptionSetType == OptionSetType.Picklist)
                    {
                        OptionSetMetadata optionSetMetadata = (OptionSetMetadata)optionSet;
                        int optionCount = 1;
                        foreach (OptionMetadata option in optionSetMetadata.Options)
                        {
                            var desc = label;
                            var value = option.Value.Value.ToString(CultureInfo.InvariantCulture);
                            var label2 = optionCount + "_" + ConvertNameAsVariable(desc);

                            string optionSetEnumCode = OptionSetEnumCodeTemplate;
                            optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Description]", TransformToSymmary(desc));
                            optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Label]", label2);
                            optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Value]", value);
                            optionSetEnums += optionSetEnumCode;
                            optionCount++;
                        }
                    }
                    else if ((OptionSetType)optionSet.OptionSetType == OptionSetType.Boolean)
                    {
                        BooleanOptionSetMetadata optionSetMetadata = (BooleanOptionSetMetadata)optionSet;

                        string optionSetEnumCode = OptionSetEnumCodeTemplate;
                        optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Description]", optionSetMetadata.TrueOption.Label.UserLocalizedLabel.Label);
                        optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Label]", "TrueOption");
                        optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Value]", optionSetMetadata.TrueOption.Value.ToString());
                        optionSetEnums += optionSetEnumCode;

                        optionSetEnumCode = OptionSetEnumCodeTemplate;
                        optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Description]", optionSetMetadata.FalseOption.Label.UserLocalizedLabel.Label);
                        optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Label]", "FalseOption");
                        optionSetEnumCode = optionSetEnumCode.Replace("[@Option.Value]", optionSetMetadata.FalseOption.Value.ToString());
                        optionSetEnums += optionSetEnumCode;
                    }
                    optionSetCode = optionSetCode.Replace("[@OptionSet.Values]", optionSetEnums);
                }

                content += optionSetCode;
            }

            string classContent = GlobalOptionSetClassCodeTemplate;
            classContent = classContent.Replace("[@DefaultNamespace]", DefaultNamespace);
            classContent = classContent.Replace("[@OptionSetDefinition]", content);

            string fileName = Path.Combine(TargetPath, "GlobalOptionSet.generated.cs");
            File.WriteAllText(fileName, classContent);
        }

        private static string TransformToSymmary(string desc)
        {
            return desc?.Replace("\n", "\n /// ");
        }

        #endregion

        #region Get Option Set Label
        /// <summary>
        /// Gets the option set label.
        /// </summary>
        /// <param name="optionSet">The option set.</param>
        /// <returns></returns>
        private static string GetOptionSetLabel(OptionSetMetadataBase optionSet)
        {
            if (optionSet.DisplayName != null && optionSet.DisplayName.UserLocalizedLabel != null && optionSet.DisplayName.UserLocalizedLabel.Label != null)
            {
                return optionSet.DisplayName.UserLocalizedLabel.Label;
            }
            return optionSet.Name;
        } 
        #endregion

        #region Load Template Code
        /// <summary>
        /// Load template code file to string
        /// </summary>
        /// <param name="templateName">Template name</param>
        /// <returns>Template code</returns>
        private static string LoadTemplateCode(string templateName)
        {
            string fileName = string.Format(@"Templates\{0}.txt", templateName);
            string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string fullPath = Path.Combine(exeDirectory, fileName);
            return File.ReadAllText(fullPath);
        }
        #endregion

        #region Convert Name As Variable
        /// <summary>
        /// Converts the name as variable.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        private static string ConvertNameAsVariable(string name)
        {
            const char separator = '_';

            // Special cases
            // Ex 1 : '1. Prospect' => 'Prospect'
            if (Regex.IsMatch(name, "[0-9]+[.] [a-z A-Z]+"))
            {
                int startPos = name.IndexOf('.') + 2;
                name = name.Substring(startPos);
            }

            // Ex 2 : '1.Prospect' => 'Prospect'
            // Ex 3 : '10. 500 à 999 personnes' => '500 à 999 personnes'
            if (Regex.IsMatch(name, "[0-9]+[.][a-z A-Z]+"))
            {
                int startPos = name.IndexOf('.') + 1;
                name = name.Substring(startPos);
            }

            // Ex 4 : '10000 - 25000' => '10000 à 25000'
            if (Regex.IsMatch(name, "[0-9]+ [-] [0-9]+"))
            {
                name = name.Replace(" - ", " à ");
            }

            // Ex 5 : 'FO - FORCE OUVRIERE' => 'FO ou FORCE OUVRIERE'
            if (Regex.IsMatch(name, "[a-z A-Z]+ [-] [a-z A-Z]+"))
            {
                name = name.Replace(" - ", " ou ");
            }

            // TODO : Find a better solution
            name = name.Replace("de – de", "de moins de");
            name = name.Replace("de + de", "de plus de");

            StringBuilder sb = new StringBuilder();
            string st = name.Normalize(NormalizationForm.FormD);

            foreach (char t in st)
            {
                UnicodeCategory uc = CharUnicodeInfo.GetUnicodeCategory(t);
                switch (uc)
                {
                    case UnicodeCategory.UppercaseLetter:
                    case UnicodeCategory.LowercaseLetter:
                    case UnicodeCategory.DecimalDigitNumber:
                        if (sb.ToString().LastOrDefault() == separator)
                        {
                            sb.Append(t.ToString().ToUpper());
                        }
                        else
                        {
                            sb.Append(t);
                        }
                        break;
                    case UnicodeCategory.SpaceSeparator:
                        if (sb.ToString().LastOrDefault() != separator)
                        {
                            sb.Append(separator);
                        }
                        break;
                }
            }

            string variable = sb.ToString().Normalize(NormalizationForm.FormC);
            variable = variable.Replace(separator.ToString(), string.Empty);
            return variable;
        }
        #endregion


    }
}
