using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Resources;
using Mono.Cecil;
using Mono.Cecil.Cil;

internal static class Program
{
    private static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        if (args.Length < 1)
        {
            Console.Error.WriteLine("Usage: InspectRecoValidation <assembly> [search-text]");
            return 2;
        }

        string assemblyPath = Path.GetFullPath(args[0]);
        string search = args.Length > 1 ? args[1] : "费用类型";
        bool methodMode = String.Equals(search, "--method", StringComparison.OrdinalIgnoreCase);
        string methodSearch = methodMode && args.Length > 2 ? args[2] : "";
        bool typeMode = String.Equals(search, "--type", StringComparison.OrdinalIgnoreCase);
        string typeSearch = typeMode && args.Length > 2 ? args[2] : "";
        bool skipSearch = String.Equals(search, "--raw", StringComparison.OrdinalIgnoreCase);
        long targetRawOffset = -1;
        if (args.Length > 2 && !methodMode && !typeMode)
        {
            string rawText = args[2].StartsWith("0x", StringComparison.OrdinalIgnoreCase)
                ? args[2].Substring(2)
                : args[2];
            targetRawOffset = Int64.Parse(rawText, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        }
        DefaultAssemblyResolver resolver = new DefaultAssemblyResolver();
        resolver.AddSearchDirectory(Path.GetDirectoryName(assemblyPath));

        using (AssemblyDefinition assembly = AssemblyDefinition.ReadAssembly(
            assemblyPath,
            new ReaderParameters { AssemblyResolver = resolver, ReadSymbols = false }))
        {
            int matches = 0;
            foreach (ModuleDefinition module in assembly.Modules)
            {
                foreach (TypeDefinition type in module.Types)
                {
                    if (typeMode)
                    {
                        matches += DumpTypes(type, typeSearch);
                    }
                    else if (methodMode)
                    {
                        matches += DumpMethods(type, methodSearch);
                    }
                    else if (!skipSearch)
                    {
                        matches += ScanType(type, search);
                    }
                }
            }

            if (!skipSearch && !methodMode && !typeMode)
            {
                matches += ScanResources(assembly, search);
            }
            if (targetRawOffset >= 0)
            {
                DumpMethodAtRawOffset(assembly, assemblyPath, targetRawOffset);
            }
            Console.WriteLine("MATCHES=" + matches.ToString(CultureInfo.InvariantCulture));
        }

        return 0;
    }

    private static int DumpTypes(TypeDefinition type, string search)
    {
        int matches = 0;
        if (type.FullName.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
        {
            Console.WriteLine("TYPE=" + type.FullName);
            Console.WriteLine("BASE=" + (type.BaseType == null ? "" : type.BaseType.FullName));
            foreach (FieldDefinition field in type.Fields)
            {
                Console.WriteLine(
                    "FIELD=" + field.Attributes
                    + " " + field.FieldType.FullName
                    + " " + field.Name);
            }
            foreach (PropertyDefinition property in type.Properties)
            {
                Console.WriteLine(
                    "PROPERTY=" + property.PropertyType.FullName
                    + " " + property.Name);
            }
            foreach (MethodDefinition method in type.Methods)
            {
                Console.WriteLine(
                    "METHOD=" + method.Attributes
                    + " " + method.ReturnType.FullName
                    + " " + method.Name
                    + "(" + String.Join(",", method.Parameters.Select(x => x.ParameterType.FullName).ToArray()) + ")");
            }
            Console.WriteLine();
            matches++;
        }

        foreach (TypeDefinition nested in type.NestedTypes)
        {
            matches += DumpTypes(nested, search);
        }

        return matches;
    }

    private static int DumpMethods(TypeDefinition type, string search)
    {
        int matches = 0;
        foreach (MethodDefinition method in type.Methods)
        {
            if (method.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) < 0)
            {
                continue;
            }

            Console.WriteLine("TYPE=" + type.FullName);
            Console.WriteLine("METHOD=" + method.FullName);
            Console.WriteLine("METHOD_TOKEN=0x" + method.MetadataToken.ToInt32().ToString("X8", CultureInfo.InvariantCulture));
            Console.WriteLine("RVA=0x" + method.RVA.ToString("X", CultureInfo.InvariantCulture));
            Console.WriteLine("HAS_BODY=" + method.HasBody.ToString(CultureInfo.InvariantCulture));
            if (method.HasBody)
            {
                Console.WriteLine("CODE_SIZE=" + method.Body.CodeSize.ToString(CultureInfo.InvariantCulture));
                foreach (Instruction instruction in method.Body.Instructions)
                {
                    Console.WriteLine(
                        "IL_" + instruction.Offset.ToString("X4", CultureInfo.InvariantCulture)
                        + ": " + instruction.OpCode.Name
                        + (instruction.Operand == null ? "" : " " + instruction.Operand));
                }
            }
            Console.WriteLine();
            matches++;
        }

        foreach (TypeDefinition nested in type.NestedTypes)
        {
            matches += DumpMethods(nested, search);
        }

        return matches;
    }

    private static int ScanType(TypeDefinition type, string search)
    {
        int matches = 0;
        foreach (MethodDefinition method in type.Methods)
        {
            if (!method.HasBody)
            {
                continue;
            }

            foreach (Instruction instruction in method.Body.Instructions)
            {
                if (instruction.OpCode != OpCodes.Ldstr)
                {
                    continue;
                }

                string text = instruction.Operand as string ?? "";
                if (text.IndexOf(search, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                Console.WriteLine("TYPE=" + type.FullName);
                Console.WriteLine("METHOD=" + method.FullName);
                Console.WriteLine("METHOD_TOKEN=0x" + method.MetadataToken.ToInt32().ToString("X8", CultureInfo.InvariantCulture));
                Console.WriteLine("IL_OFFSET=0x" + instruction.Offset.ToString("X", CultureInfo.InvariantCulture));
                Console.WriteLine("STRING=" + text.Replace("\r", " ").Replace("\n", " "));
                Console.WriteLine();
                matches++;
            }
        }

        foreach (TypeDefinition nested in type.NestedTypes)
        {
            matches += ScanType(nested, search);
        }

        return matches;
    }

    private static void DumpMethodAtRawOffset(AssemblyDefinition assembly, string assemblyPath, long targetRawOffset)
    {
        byte[] image = File.ReadAllBytes(assemblyPath);
        foreach (ModuleDefinition module in assembly.Modules)
        {
            foreach (TypeDefinition type in module.Types)
            {
                if (DumpTypeMethodAtRawOffset(type, image, targetRawOffset))
                {
                    return;
                }
            }
        }

        Console.WriteLine("RAW_METHOD_NOT_FOUND=0x" + targetRawOffset.ToString("X", CultureInfo.InvariantCulture));
        int targetRva = RawToRva(image, checked((int)targetRawOffset));
        Console.WriteLine("RAW_TARGET_RVA=0x" + targetRva.ToString("X", CultureInfo.InvariantCulture));
        foreach (MethodDefinition method in AllMethods(assembly)
            .Where(x => x.RVA != 0)
            .OrderBy(x => Math.Abs((long)x.RVA - targetRva))
            .Take(12))
        {
            Console.WriteLine(
                "NEAR=0x" + method.RVA.ToString("X", CultureInfo.InvariantCulture)
                + " " + method.FullName);
        }
    }

    private static bool DumpTypeMethodAtRawOffset(TypeDefinition type, byte[] image, long targetRawOffset)
    {
        foreach (MethodDefinition method in type.Methods)
        {
            if (!method.HasBody || method.RVA == 0)
            {
                continue;
            }

            int methodRaw = RvaToRaw(image, method.RVA);
            if (methodRaw < 0)
            {
                continue;
            }

            int first = image[methodRaw];
            int headerSize;
            int codeSize;
            if ((first & 3) == 2)
            {
                headerSize = 1;
                codeSize = first >> 2;
            }
            else
            {
                ushort flags = BitConverter.ToUInt16(image, methodRaw);
                headerSize = ((flags >> 12) & 0x0f) * 4;
                codeSize = BitConverter.ToInt32(image, methodRaw + 4);
            }

            long codeRaw = methodRaw + headerSize;
            if (targetRawOffset < codeRaw || targetRawOffset >= codeRaw + codeSize)
            {
                continue;
            }

            int targetIlOffset = checked((int)(targetRawOffset - codeRaw));
            Console.WriteLine("RAW_TYPE=" + type.FullName);
            Console.WriteLine("RAW_METHOD=" + method.FullName);
            Console.WriteLine("RAW_METHOD_TOKEN=0x" + method.MetadataToken.ToInt32().ToString("X8", CultureInfo.InvariantCulture));
            Console.WriteLine("RAW_TARGET_IL_OFFSET=0x" + targetIlOffset.ToString("X", CultureInfo.InvariantCulture));
            foreach (Instruction instruction in method.Body.Instructions)
            {
                if (instruction.Offset < targetIlOffset - 80 || instruction.Offset > targetIlOffset + 120)
                {
                    continue;
                }

                Console.WriteLine(
                    "IL_" + instruction.Offset.ToString("X4", CultureInfo.InvariantCulture)
                    + ": " + instruction.OpCode.Name
                    + (instruction.Operand == null ? "" : " " + instruction.Operand));
            }

            return true;
        }

        foreach (TypeDefinition nested in type.NestedTypes)
        {
            if (DumpTypeMethodAtRawOffset(nested, image, targetRawOffset))
            {
                return true;
            }
        }

        return false;
    }

    private static int RvaToRaw(byte[] image, int rva)
    {
        int pe = BitConverter.ToInt32(image, 0x3c);
        int sectionCount = BitConverter.ToUInt16(image, pe + 6);
        int optionalSize = BitConverter.ToUInt16(image, pe + 20);
        int sectionHeaders = pe + 24 + optionalSize;
        for (int i = 0; i < sectionCount; i++)
        {
            int header = sectionHeaders + i * 40;
            uint virtualSize = BitConverter.ToUInt32(image, header + 8);
            uint virtualAddress = BitConverter.ToUInt32(image, header + 12);
            uint rawSize = BitConverter.ToUInt32(image, header + 16);
            uint rawAddress = BitConverter.ToUInt32(image, header + 20);
            uint span = Math.Max(virtualSize, rawSize);
            if ((uint)rva >= virtualAddress && (uint)rva < virtualAddress + span)
            {
                return checked((int)(rawAddress + (uint)rva - virtualAddress));
            }
        }

        return -1;
    }

    private static int RawToRva(byte[] image, int raw)
    {
        int pe = BitConverter.ToInt32(image, 0x3c);
        int sectionCount = BitConverter.ToUInt16(image, pe + 6);
        int optionalSize = BitConverter.ToUInt16(image, pe + 20);
        int sectionHeaders = pe + 24 + optionalSize;
        for (int i = 0; i < sectionCount; i++)
        {
            int header = sectionHeaders + i * 40;
            uint virtualAddress = BitConverter.ToUInt32(image, header + 12);
            uint rawSize = BitConverter.ToUInt32(image, header + 16);
            uint rawAddress = BitConverter.ToUInt32(image, header + 20);
            if ((uint)raw >= rawAddress && (uint)raw < rawAddress + rawSize)
            {
                return checked((int)(virtualAddress + (uint)raw - rawAddress));
            }
        }

        return -1;
    }

    private static IEnumerable<MethodDefinition> AllMethods(AssemblyDefinition assembly)
    {
        foreach (ModuleDefinition module in assembly.Modules)
        {
            foreach (TypeDefinition type in module.Types)
            {
                foreach (MethodDefinition method in AllMethods(type))
                {
                    yield return method;
                }
            }
        }
    }

    private static IEnumerable<MethodDefinition> AllMethods(TypeDefinition type)
    {
        foreach (MethodDefinition method in type.Methods)
        {
            yield return method;
        }

        foreach (TypeDefinition nested in type.NestedTypes)
        {
            foreach (MethodDefinition method in AllMethods(nested))
            {
                yield return method;
            }
        }
    }

    private static int ScanResources(AssemblyDefinition assembly, string search)
    {
        int matches = 0;
        foreach (Resource resource in assembly.MainModule.Resources)
        {
            EmbeddedResource embedded = resource as EmbeddedResource;
            if (embedded == null)
            {
                continue;
            }

            try
            {
                using (Stream stream = embedded.GetResourceStream())
                using (ResourceReader reader = new ResourceReader(stream))
                {
                    System.Collections.IDictionaryEnumerator values = reader.GetEnumerator();
                    while (values.MoveNext())
                    {
                        string text = values.Value as string;
                        if (text == null || text.IndexOf(search, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            continue;
                        }

                        Console.WriteLine("RESOURCE=" + resource.Name);
                        Console.WriteLine("KEY=" + Convert.ToString(values.Key, CultureInfo.InvariantCulture));
                        Console.WriteLine("STRING=" + text.Replace("\r", " ").Replace("\n", " "));
                        Console.WriteLine();
                        matches++;
                    }
                }
            }
            catch
            {
            }
        }

        return matches;
    }
}
