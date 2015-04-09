namespace ActiveXScriptLib.Extensions
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

         ProcessOptions(builder, options);

         return builder;
      }

      private static void ProcessOptions(ActiveScriptEngineBuilder builder, IniOptions options)
      {
         string iniDirectory = Path.GetDirectoryName(options._iniFilePath);

         var parser = new FileIniDataParser();

         IniData data = parser.ReadFile(options._iniFilePath);

         KeyDataCollection sectionData = data[options._sectionName];

         foreach (KeyData kvp in sectionData)
         {
            // E.g. "Script1.vbs" and "Namespace".
            string[] values = kvp.Value.Replace("\"", "").Split(',');

            string scriptPath = values[0].Trim();

            foreach (var tokens in options._tokens)
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
