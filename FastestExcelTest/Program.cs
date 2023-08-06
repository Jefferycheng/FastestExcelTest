// See https://aka.ms/new-console-template for more information

using BenchmarkDotNet.Running;
using FastestExcelTest;

Console.WriteLine("Hello, World!");

var summary = BenchmarkRunner.Run<CompareExcel>();

Console.WriteLine("Bye Bye !");