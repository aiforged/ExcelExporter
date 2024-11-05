using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

using LiquidTechnologies.XmlObjects;

namespace ExcelExporter.Helpers
{
    public static partial class Tools
    {
        public static void SerializeObjectToFile<T>(T obj, string filePath)
        {
            File.WriteAllText(filePath, ToXml(obj));
        }

        public static void SerializeObjectToJson<T>(T obj, string filePath)
        {
            File.WriteAllText(filePath, System.Text.Json.JsonSerializer.Serialize(obj, new System.Text.Json.JsonSerializerOptions()
            {
                AllowTrailingCommas = true,
                DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
                WriteIndented = true,
            }));
        }

        public static string ToXml<T>(T obj)
        {
            string result = null;

            var settings = new XmlWriterSettings
            {
                Indent = true,
                OmitXmlDeclaration = true,
            };

            var ns = new XmlSerializerNamespaces(new[] { XmlQualifiedName.Empty });

            // Create an instance of XmlSerializer specifying type of object to serialize
            XmlSerializer serializer = new XmlSerializer(typeof(T), overrides: null, extraTypes: null, new XmlRootAttribute(typeof(T).Name), defaultNamespace: "http://www.acord.org/schema/data/draft/ReusableDataComponents/1");

            // Open a file stream to write the file
            using (StringWriter stream = new StringWriter())
            {
                using (var writer = XmlWriter.Create(stream, settings))
                {
                    // Serialize the object to XML and write it to the file stream
                    serializer.Serialize(writer, obj, ns);
                    result = stream.ToString();
                }
            }

            // Replace all nullable fields, other solution would be to use add PropSpecified property for all properties that are not strings
            result = Regex.Replace(result, "\\s*<\\w+ [a-z\\d]*:nil=\\\"true\\\".*\\/>", string.Empty);

            return result;
        }

        public static string StringToGuidFormat(string input)
        {
            if (input == null)
                throw new ArgumentNullException(nameof(input));

            // Remove any hyphens that might be in the input string
            string sanitizedInput = input.Replace("-", string.Empty);

            // Ensure the string contains only hexadecimal characters
            if (!System.Text.RegularExpressions.Regex.IsMatch(sanitizedInput, @"\A\b[0-9a-fA-F]+\b\Z"))
                throw new ArgumentException("Input string contains non-hexadecimal characters.", nameof(input));

            // Ensure the length is less than or equal to 32 characters (for GUID format without hyphens)
            if (sanitizedInput.Length > 32)
                throw new ArgumentException("Input string is too long to be formatted as GUID.", nameof(input));

            // Pad the string with zeroes to ensure it is 32 characters long
            sanitizedInput = sanitizedInput.PadRight(32, '0');

            // Insert hyphens to format as GUID: 8-4-4-4-12
            return $"{sanitizedInput.Substring(0, 8)}-{sanitizedInput.Substring(8, 4)}-{sanitizedInput.Substring(12, 4)}-{sanitizedInput.Substring(16, 4)}-{sanitizedInput.Substring(20, 12)}";
        }

        public static string StringToPhoneFormat(string input)
        {
            if (input == null)
                throw new ArgumentNullException(nameof(input));

            // Remove any hyphens that might be in the input string
            string sanitizedInput = input.Replace("-", string.Empty);
            sanitizedInput = sanitizedInput.Replace("+", string.Empty);
            sanitizedInput = sanitizedInput.Replace(" ", string.Empty);

            // Ensure the string contains only hexadecimal characters
            if (!System.Text.RegularExpressions.Regex.IsMatch(sanitizedInput, @"[0-9]+"))
                throw new ArgumentException("Input string contains non-numeric characters.", nameof(input));

            // Ensure the length is less than or equal to 32 characters (for GUID format without hyphens)
            if (sanitizedInput.Length < 11)
                if (sanitizedInput.Length == 10)
                {
                    return $"+27-{sanitizedInput.Substring(1, 2)}-{sanitizedInput.Substring(3, 3)}-{sanitizedInput.Substring(6)}";
                }
                else
                {
                    throw new ArgumentException("Input string is too short to be formatted as phone.", nameof(input));
                }

            // Insert hyphens to format as GUID: 8-4-4-4-12
            return $"+{sanitizedInput.Substring(0, 2)}-{sanitizedInput.Substring(2, 2)}-{sanitizedInput.Substring(4, 3)}-{sanitizedInput.Substring(7)}";
        }

        public static List<string> GetEnumValues<T>()
        {
            return Enum.GetValues(typeof(T)).Cast<T>().Select(c => $"{c}").ToList();
        }

        public static string GetBestMatchingEnumValue<T>(string enumValue, float minConf, string fallback)
        {
            string returnData = $"{default(T)}";
            float? maxConf = 0.0F;

            List<string> enumValues = GetEnumValues<T>();

            foreach (string enumV in enumValues)
            {
                float? levConf = LarcAI.Utilities.LevenshteinConfidence(enumV, LarcAI.Utilities.LevenshteinDistance(enumV, enumValue), enumValue);

                if (levConf > maxConf)
                {
                    maxConf = levConf;
                    returnData = enumV;
                }
            }

            if (maxConf < minConf)
            {
                returnData = fallback;
            }

            return returnData;
        }

        public static T GetBestMatchingEnum<T>(string enumValue, float minConf = 0.5F, string fallback = "Other")
            where T: struct
        {
            if (fallback == null)
            {
                fallback = $"{default(T)}";
            }

            return LarcAI.Utilities.GetEnum<T>(GetBestMatchingEnumValue<T>(enumValue, minConf, fallback));
        }

        public static void AssignByMembers<T>(object src, T dst, bool propertiesOnly = false, params string[] excludeMemberNames)
        {
            try
            {
                if (dst is null || src is null)
                {
                    return;
                }

                Type type = dst.GetType();
                Type type2 = src!.GetType();
                PropertyInfo[] properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.SetProperty);

                foreach (PropertyInfo propertyInfo in properties)
                {
                    try
                    {
                        PropertyInfo property = type2.GetProperty(propertyInfo.Name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.GetProperty);

                        if (property != null && property.PropertyType == propertyInfo.PropertyType && propertyInfo.GetSetMethod() != null && !excludeMemberNames.Contains(property.Name))
                        {
                            object value = property.GetValue(src, null);
                            propertyInfo.SetValue(dst, value, null);
                        }
                    }
                    catch
                    {
                    }
                }

                if (propertiesOnly) return;
                FieldInfo[] fields = type.GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.SetField);

                foreach (FieldInfo fieldInfo in fields)
                {
                    try
                    {
                        FieldInfo field = type2.GetField(fieldInfo.Name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.GetField);

                        if (field != null && field.FieldType == fieldInfo.FieldType && !excludeMemberNames.Contains(field.Name))
                        {
                            object value = field.GetValue(src);
                            fieldInfo.SetValue(dst, value);
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        public static Stack<T> CloneStack<T>(Stack<T> stack)
        {
            return new Stack<T>(stack.Reverse());
        }

        public static object[,] GetArraySection(object[,] array, int startRow, int numberOfRows, int startCol, int numberOfCols)
        {
            // Initialize the new 2D array 
            object[,] newArray = new object[numberOfRows, numberOfCols];

            // Copy elements from the original array to the new array 
            for (int i = 0; i < numberOfRows; i++)
            {
                for (int j = 0; j < numberOfCols; j++)
                {
                    newArray[i, j] = array[startRow + i, startCol + j];
                }
            }

            return newArray;
        }
    }
}
