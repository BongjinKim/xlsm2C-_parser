# xlsm2csharp_parser
parsing to xlsm2C#
![https://s3-us-west-2.amazonaws.com/secure.notion-static.com/7e8fdb90-7b58-41d2-82b4-84e2349ac672/Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/7e8fdb90-7b58-41d2-82b4-84e2349ac672/Untitled.png)

패턴에 따른 몬스터 출현 테이블 Parser 제작. 해당 데이터를 넣으면 VBA로 parsing하여 자동으로 C#파일로 만들어준다. xml → C# 혹은 json → C# 의 작업이 필요 없게 된다.

```visual-basic
Sub ExportPetDataTable()

    Dim 탭, 쉼표 As String
    탭 = Chr(9)
    쉼표 = ","

    Open ThisWorkbook.Path & "\..\source\Shooting\Assets\Game\Scripts\Data\SpawnPatternDataLoader.cs" For Output As #1
    Print #1, "using UnityEngine;"
    Print #1, "using System.Collections.Generic;"
    Print #1,
    Print #1, "public class SpawnPatternDataLoader : IDataLoader"
    Print #1, "{"
    Print #1, 탭 & "List<MonsterSpawnPatternData> spawnPatternList = new List<MonsterSpawnPatternData>();"
    Print #1,
    Print #1, 탭 & "public void LoadData()"
    Print #1, 탭 & "{"
    
    Dim MaxRow, MaxCol
    MaxRow = Range("A4").value * 7 + 6
    MaxCol = Range("B4").value + 5 + 4
    
    Set Pattern = Range(Cells(7, 3), Cells(MaxRow, 3))
    
    For Each p In Pattern
        If Not IsEmpty(p.value) Then
            Print #1, 탭 & 탭 & "spawnPatternList.Add(new MonsterSpawnPatternData()"
            Print #1, 탭 & 탭 & "{"
            
            Print #1, 탭 & 탭 & 탭 & "SpawnMonsterDataArray = new SpawnMonsterData[]"
            Print #1, 탭 & 탭 & 탭 & "{"
                            
            Set OnePattern = Range(Cells(p.Row, 5), Cells(p.Row, MaxCol))
            
            For Each op In OnePattern
                If Not IsEmpty(op.value) Then
                            
                Print #1, 탭 & 탭 & 탭 & 탭 & "new SpawnMonsterData()"
                Print #1, 탭 & 탭 & 탭 & 탭 & "{"
                Print #1, 탭 & 탭 & 탭 & 탭 & 탭 & "Monster = " & FindValueMonster(Cells(op.Row, op.Column)) & 쉼표; ""
                Print #1, 탭 & 탭 & 탭 & 탭 & 탭 & "Level = " & Cells(op.Row + 1, op.Column) & 쉼표
                Print #1, 탭 & 탭 & 탭 & 탭 & 탭 & "Count = " & Cells(op.Row + 2, op.Column) & 쉼표
								
								"... 중략 ..."

End Sub
```
