using System;
using System.IO;
using System.Linq;
using dnlib.DotNet;
using dnlib.DotNet.Emit;

internal static class InjectRecoExpandPanel
{
    private static int Main(string[] args)
    {
        if (args.Length != 1)
        {
            Console.Error.WriteLine("Usage: InjectRecoExpandPanel <RejjNet2020.exe>");
            return 2;
        }

        string exePath = Path.GetFullPath(args[0]);
        string dir = Path.GetDirectoryName(exePath);
        string backup = Path.Combine(dir, "RejjNet2020.exe.pre-autoload.bak");

        if (!File.Exists(backup))
        {
            File.Copy(exePath, backup);
        }

        ModuleDefMD module = ModuleDefMD.Load(exePath);
        TypeDef mainForm = module.Find("RecoNet.RecoMainForm", false);
        if (mainForm == null)
        {
            Console.Error.WriteLine("RecoNet.RecoMainForm not found.");
            return 3;
        }

        AssemblyRef expandAssembly = module.GetAssemblyRefs()
            .FirstOrDefault(a => a.Name == "RecoExpandPanel");
        if (expandAssembly == null)
        {
            expandAssembly = new AssemblyRefUser("RecoExpandPanel", new Version(0, 0, 0, 0));
        }

        ITypeDefOrRef formType = module.Import(typeof(System.Windows.Forms.Form));
        TypeRef formPanelType = new TypeRefUser(module, "RecoNet", "FormPanel", expandAssembly);
        MemberRef ctor = new MemberRefUser(
            module,
            ".ctor",
            MethodSig.CreateInstance(module.CorLibTypes.Void, formType.ToTypeSig()),
            formPanelType);

        int injected = 0;
        foreach (MethodDef method in mainForm.Methods.Where(m => m.Name == ".ctor" && m.HasBody))
        {
            if (AlreadyInjected(method))
            {
                continue;
            }

            method.Body.SimplifyBranches();
            for (int i = method.Body.Instructions.Count - 1; i >= 0; i--)
            {
                if (method.Body.Instructions[i].OpCode == OpCodes.Ret)
                {
                    method.Body.Instructions.Insert(i++, Instruction.Create(OpCodes.Ldarg_0));
                    method.Body.Instructions.Insert(i++, Instruction.Create(OpCodes.Newobj, ctor));
                    method.Body.Instructions.Insert(i++, Instruction.Create(OpCodes.Pop));
                    injected++;
                }
            }
            method.Body.OptimizeBranches();
        }

        if (injected == 0)
        {
            Console.Error.WriteLine("No constructor return point injected.");
            return 4;
        }

        string patched = exePath + ".patched";
        module.Write(patched);
        File.Copy(patched, exePath, true);
        File.Delete(patched);
        Console.WriteLine("Injected FormPanel autoload at " + injected + " constructor return point(s).");
        Console.WriteLine("Backup: " + backup);
        return 0;
    }

    private static bool AlreadyInjected(MethodDef method)
    {
        return method.Body.Instructions.Any(i =>
        {
            IMethod methodRef = i.Operand as IMethod;
            return i.OpCode == OpCodes.Newobj &&
                   methodRef != null &&
                   methodRef.DeclaringType.FullName == "RecoNet.FormPanel";
        });
    }
}
