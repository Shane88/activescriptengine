namespace ActiveXScriptLib.Extensions.Builder
{
   using System;
   using System.IO;
   using IniParser;
   using IniParser.Model;

   public static class ActiveScriptEngineIniFileExtensions
   {
      public static ActiveScriptEngineBuilder AddFilesFromIni(this ActiveScriptEngineBuilder builder, Action<IniOptions> optionConfig)
      {
         IniOptions options = new IniOptions();

         optionConfig(options);

         KeyDataCollection sectionData = ParseAndValidateOptions(options);

         ProcessOptions(builder, options, sectionData);

         return builder;
      }

      private static KeyDataCollection ParseAndValidateOptions(IniOptions options)
      {
         var parser = new FileIniDataParser();

         IniData data = null;

         try
         {
            data = parser.ReadFile(options.ConfiguredIniFilePath);
         }
         catch (Exception ex)
         {
            string errorMessage = string.Format(
               "An error occurred when trying to Read Ini file {0}. Check the InnerException for more details.",
               options.ConfiguredIniFilePath);

            throw new InvalidOperationException(errorMessage, ex);
         }

         if (!data.Sections.ContainsSection(options.ConfiguredSectionName))
         {
            string errorMessage = string.Format(
               "An error occurred when trying to read the IniFile {0}. No such section exists called {1}.",
               options.ConfiguredIniFilePath,
               options.ConfiguredSectionName);

            throw new InvalidOperationException(errorMessage);
         }

         KeyDataCollection sectionData = data[options.ConfiguredSectionName];

         return sectionData;
      }

      private static void ProcessOptions(ActiveScriptEngineBuilder builder, IniOptions options, KeyDataCollection sectionData)
      {
         string iniDirectory = Path.GetDirectoryName(options.ConfiguredIniFilePath);

         foreach (KeyData kvp in sectionData)
         {
            // E.g. "Script1.vbs" and "Namespace".
            string[] values = kvp.Value.Replace("\"", "").Split(',');

            string scriptPath = values[0].Trim();

            foreach (var tokens in options.ConfiguredTokens)
            {
               scriptPath = scriptPath.Replace(tokens.Key, tokens.Value);
            }

            string namespaceName = null;

            if (values.Length > 1 && values[1] != null)
            {
               namespaceName = values[1].Trim();
            }

            builder.AddCodeFiles(scriptPath, namespaceName, iniDirectory);
         }
      }
   }
}
