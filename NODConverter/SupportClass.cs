//
// In order to convert some functionality to Visual C#, the Java Language Conversion Assistant
// creates "support classes" that duplicate the original functionality.  
//
// Support classes replicate the functionality of the original code, but in some cases they are 
// substantially different architecturally. Although every effort is made to preserve the 
// original architecture of the application in the converted project, the user should be aware that 
// the primary goal of these support classes is to replicate functionality, and that at times 
// the architecture of the resulting solution may differ somewhat.
//

using System;

/// <summary>
/// Contains conversion support elements such as classes, interfaces and static methods.
/// </summary>
public class SupportClass
{
    /// <summary>
    /// SupportClass for the HashSet class.
    /// </summary>
    [Serializable]
    public class HashSetSupport : System.Collections.ArrayList, SetSupport
    {
        public HashSetSupport()
            : base()
        {
        }

        public HashSetSupport(System.Collections.ICollection c)
        {
            this.AddAll(c);
        }

        public HashSetSupport(int capacity)
            : base(capacity)
        {
        }

        /// <summary>
        /// Adds a new element to the ArrayList if it is not already present.
        /// </summary>		
        /// <param name="obj">Element to insert to the ArrayList.</param>
        /// <returns>Returns true if the new element was inserted, false otherwise.</returns>
        new public virtual bool Add(System.Object obj)
        {
            bool inserted;

            if ((inserted = this.Contains(obj)) == false)
            {
                base.Add(obj);
            }

            return !inserted;
        }

        /// <summary>
        /// Adds all the elements of the specified collection that are not present to the list.
        /// </summary>
        /// <param name="c">Collection where the new elements will be added</param>
        /// <returns>Returns true if at least one element was added, false otherwise.</returns>
        public bool AddAll(System.Collections.ICollection c)
        {
            System.Collections.IEnumerator e = new System.Collections.ArrayList(c).GetEnumerator();
            bool added = false;

            while (e.MoveNext() == true)
            {
                if (this.Add(e.Current) == true)
                    added = true;
            }

            return added;
        }

        /// <summary>
        /// Returns a copy of the HashSet instance.
        /// </summary>		
        /// <returns>Returns a shallow copy of the current HashSet.</returns>
        public override System.Object Clone()
        {
            return base.MemberwiseClone();
        }
    }


    /*******************************/
    /// <summary>
    /// Represents a collection ob objects that contains no duplicate elements.
    /// </summary>	
    public interface SetSupport : System.Collections.ICollection, System.Collections.IList
    {
        /// <summary>
        /// Adds a new element to the Collection if it is not already present.
        /// </summary>
        /// <param name="obj">The object to add to the collection.</param>
        /// <returns>Returns true if the object was added to the collection, otherwise false.</returns>
        new bool Add(System.Object obj);

        /// <summary>
        /// Adds all the elements of the specified collection to the Set.
        /// </summary>
        /// <param name="c">Collection of objects to add.</param>
        /// <returns>true</returns>
        bool AddAll(System.Collections.ICollection c);
    }


    /*******************************/
    /// <summary>
    /// Provides functionality not found in .NET map-related interfaces.
    /// </summary>
    public class MapSupport
    {
        /// <summary>
        /// Determines whether the SortedList contains a specific value.
        /// </summary>
        /// <param name="d">The dictionary to check for the value.</param>
        /// <param name="obj">The object to locate in the SortedList.</param>
        /// <returns>Returns true if the value is contained in the SortedList, false otherwise.</returns>
        public static bool ContainsValue(System.Collections.IDictionary d, System.Object obj)
        {
            bool contained = false;
            System.Type type = d.GetType();

            //Classes that implement the SortedList class
            if (type == System.Type.GetType("System.Collections.SortedList"))
            {
                contained = (bool)((System.Collections.SortedList)d).ContainsValue(obj);
            }
            //Classes that implement the Hashtable class
            else if (type == System.Type.GetType("System.Collections.Hashtable"))
            {
                contained = (bool)((System.Collections.Hashtable)d).ContainsValue(obj);
            }
            else
            {
                //Reflection. Invoke "containsValue" method for proprietary classes
                try
                {
                    System.Reflection.MethodInfo method = type.GetMethod("containsValue");
                    contained = (bool)method.Invoke(d, new Object[] { obj });
                }
                catch (System.Reflection.TargetInvocationException e)
                {
                    throw e;
                }
                catch (System.Exception e)
                {
                    throw e;
                }
            }

            return contained;
        }


        /// <summary>
        /// Determines whether the NameValueCollection contains a specific value.
        /// </summary>
        /// <param name="d">The dictionary to check for the value.</param>
        /// <param name="obj">The object to locate in the SortedList.</param>
        /// <returns>Returns true if the value is contained in the NameValueCollection, false otherwise.</returns>
        public static bool ContainsValue(System.Collections.Specialized.NameValueCollection d, System.Object obj)
        {
            bool contained = false;
            System.Type type = d.GetType();

            for (int i = 0; i < d.Count && !contained; i++)
            {
                System.String[] values = d.GetValues(i);
                if (values != null)
                {
                    foreach (System.String val in values)
                    {
                        if (val.Equals(obj))
                        {
                            contained = true;
                            break;
                        }
                    }
                }
            }
            return contained;
        }

        /// <summary>
        /// Copies all the elements of d to target.
        /// </summary>
        /// <param name="target">Collection where d elements will be copied.</param>
        /// <param name="d">Elements to copy to the target collection.</param>
        public static void PutAll(System.Collections.IDictionary target, System.Collections.IDictionary d)
        {
            if (d != null)
            {
                System.Collections.ArrayList keys = new System.Collections.ArrayList(d.Keys);
                System.Collections.ArrayList values = new System.Collections.ArrayList(d.Values);

                for (int i = 0; i < keys.Count; i++)
                    target[keys[i]] = values[i];
            }
        }

        /// <summary>
        /// Returns a portion of the list whose keys are less than the limit object parameter.
        /// </summary>
        /// <param name="l">The list where the portion will be extracted.</param>
        /// <param name="limit">The end element of the portion to extract.</param>
        /// <returns>The portion of the collection whose elements are less than the limit object parameter.</returns>
        public static System.Collections.SortedList HeadMap(System.Collections.SortedList l, System.Object limit)
        {
            System.Collections.Comparer comparer = System.Collections.Comparer.Default;
            System.Collections.SortedList newList = new System.Collections.SortedList();

            for (int i = 0; i < l.Count; i++)
            {
                if (comparer.Compare(l.GetKey(i), limit) >= 0)
                    break;

                newList.Add(l.GetKey(i), l[l.GetKey(i)]);
            }

            return newList;
        }

        /// <summary>
        /// Returns a portion of the list whose keys are greater that the lowerLimit parameter less than the upperLimit parameter.
        /// </summary>
        /// <param name="l">The list where the portion will be extracted.</param>
        /// <param name="limit">The start element of the portion to extract.</param>
        /// <param name="limit">The end element of the portion to extract.</param>
        /// <returns>The portion of the collection.</returns>
        public static System.Collections.SortedList SubMap(System.Collections.SortedList list, System.Object lowerLimit, System.Object upperLimit)
        {
            System.Collections.Comparer comparer = System.Collections.Comparer.Default;
            System.Collections.SortedList newList = new System.Collections.SortedList();

            if (list != null)
            {
                if ((list.Count > 0) && (!(lowerLimit.Equals(upperLimit))))
                {
                    int index = 0;
                    while (comparer.Compare(list.GetKey(index), lowerLimit) < 0)
                        index++;

                    for (; index < list.Count; index++)
                    {
                        if (comparer.Compare(list.GetKey(index), upperLimit) >= 0)
                            break;

                        newList.Add(list.GetKey(index), list[list.GetKey(index)]);
                    }
                }
            }

            return newList;
        }

        /// <summary>
        /// Returns a portion of the list whose keys are greater than the limit object parameter.
        /// </summary>
        /// <param name="l">The list where the portion will be extracted.</param>
        /// <param name="limit">The start element of the portion to extract.</param>
        /// <returns>The portion of the collection whose elements are greater than the limit object parameter.</returns>
        public static System.Collections.SortedList TailMap(System.Collections.SortedList list, System.Object limit)
        {
            System.Collections.Comparer comparer = System.Collections.Comparer.Default;
            System.Collections.SortedList newList = new System.Collections.SortedList();

            if (list != null)
            {
                if (list.Count > 0)
                {
                    int index = 0;
                    while (comparer.Compare(list.GetKey(index), limit) < 0)
                        index++;

                    for (; index < list.Count; index++)
                        newList.Add(list.GetKey(index), list[list.GetKey(index)]);
                }
            }

            return newList;
        }
    }


    /*******************************/
    /// <summary>
    /// Deserializes an object, or an entire graph of connected objects, and returns the object intance
    /// </summary>
    /// <param name="binaryReader">Reader instance used to read the object</param>
    /// <returns>The object instance</returns>
    public static System.Object Deserialize(System.IO.BinaryReader binaryReader)
    {
        System.Runtime.Serialization.Formatters.Binary.BinaryFormatter formatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
        return formatter.Deserialize(binaryReader.BaseStream);
    }

    /*******************************/
    /// <summary>
    /// Writes an object to the specified Stream
    /// </summary>
    /// <param name="stream">The target Stream</param>
    /// <param name="objectToSend">The object to be sent</param>
    public static void Serialize(System.IO.Stream stream, System.Object objectToSend)
    {
        System.Runtime.Serialization.Formatters.Binary.BinaryFormatter formatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
        formatter.Serialize(stream, objectToSend);
    }

    /// <summary>
    /// Writes an object to the specified BinaryWriter
    /// </summary>
    /// <param name="stream">The target BinaryWriter</param>
    /// <param name="objectToSend">The object to be sent</param>
    public static void Serialize(System.IO.BinaryWriter binaryWriter, System.Object objectToSend)
    {
        System.Runtime.Serialization.Formatters.Binary.BinaryFormatter formatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
        formatter.Serialize(binaryWriter.BaseStream, objectToSend);
    }

}
