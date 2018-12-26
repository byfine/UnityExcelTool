using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public enum Animal
{
    Rabbit,
    Tiger,
    Lion
}

public class ExcelToolDemo : MonoBehaviour
{
    public TSet_Example1 Table1;
    public TSet_Example2 Table2;
    
    void Start()
    {
        Debug.Log("Table1 Count: " + Table1.Count);
        foreach (var kv in Table1)
        {
            Debug.Log("Table1 Data: " + kv.Key + " : " + kv.Value);
        }


        Debug.Log("Table1 Data 1: Name:" + Table1[1].Name + ", HP:" + Table1[1].HP + ", Attack:" + Table1[1].Attack);
        
        Debug.Log("Table2: Name:" + Table2[Animal.Rabbit].Name + ", Asset:" + Table2[Animal.Rabbit].AssetName + ", Type:" + Table2[Animal.Rabbit].Type);
        
        
    }

    void Update()
    {
        
    }
}
