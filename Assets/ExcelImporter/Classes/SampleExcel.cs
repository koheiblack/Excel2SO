using UnityEngine;
using System.Collections;
using System.Collections.Generic;
public class SampleExcel : ScriptableObject
{
    public List<Param> param;

    [System.SerializableAttribute]
    public class Param
    {
        public int ID;
        public string Name;
        public string ObjectName;
        public int IntParam;
        [System.SerializableAttribute]
        public struct tagStructParam
        {
            public float FloatParam;
            public float IntParam;
        }
        public tagStructParam StructParam;
        public float[] ArrayFloat;
    }
}
