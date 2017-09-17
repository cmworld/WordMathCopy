/********************************************************************************
 *   This file is part of NRtfTree Library.
 *
 *   NRtfTree Library is free software; you can redistribute it and/or modify
 *   it under the terms of the GNU Lesser General Public License as published by
 *   the Free Software Foundation; either version 3 of the License, or
 *   (at your option) any later version.
 *
 *   NRtfTree Library is distributed in the hope that it will be useful,
 *   but WITHOUT ANY WARRANTY; without even the implied warranty of
 *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *   GNU Lesser General Public License for more details.
 *
 *   You should have received a copy of the GNU Lesser General Public License
 *   along with this program. If not, see <http://www.gnu.org/licenses/>.
 ********************************************************************************/

/********************************************************************************
 * Library:		NRtfTree
 * Version:     v0.4
 * Date:		29/06/2013
 * Copyright:   2006-2013 Salvador Gomez
 * Home Page:	http://www.sgoliver.net
 * GitHub:	    https://github.com/sgolivernet/nrtftree
 * Class:		RtfTreeNode
 * Description:	Nodo RTF de la representaci髇 en 醨bol de un documento.
 * ******************************************************************************/

using System;
using System.Text;

namespace Net.Sgoliver.NRtfTree
{
    namespace Core
    {
        /// <summary>
        /// Nodo RTF de la representaci髇 en 醨bol de un documento.
        /// </summary>
        public class RtfTreeNode
        {
            #region Atributos Privados

            /// <summary>
            /// Tipo de nodo.
            /// </summary>
            private RtfNodeType type;
            /// <summary>
            /// Palabra clave / S韒bolo de Control / Texto.
            /// </summary>
            private string key;
            /// <summary>
            /// Indica si la palabra clave o s韒bolo de Control tiene par醡etro.
            /// </summary>
            private bool hasParam;
            /// <summary>
            /// Par醡etro de la palabra clave o s韒bolo de Control.
            /// </summary>
            private int param;
            /// <summary>
            /// Nodos hijos del nodo actual.
            /// </summary>
            private RtfNodeCollection children;
            /// <summary>
            /// Nodo padre del nodo actual.
            /// </summary>
            private RtfTreeNode parent;
            /// <summary>
            /// Nodo ra韟 del documento.
            /// </summary>
            private RtfTreeNode root;
            /// <summary>
            /// 羠bol Rtf al que pertenece el nodo
            /// </summary>
            private RtfTree tree;

            #endregion

            #region Constructores P鷅licos

            /// <summary>
            /// Constructor de la clase RtfTreeNode. Crea un nodo sin inicializar.
            /// </summary>
            public RtfTreeNode()
            {
                type = RtfNodeType.None;
                key = "";
                
				/* Inicializados por defecto */
				//this.param = 0;
				//this.hasParam = false;
                //this.parent = null;
                //this.root = null;
                //this.tree = null;
                //children = null;
            }

            /// <summary>
            /// Constructor de la clase RtfTreeNode. Crea un nodo de un tipo concreto.
            /// </summary>
            /// <param name="nodeType">Tipo del nodo que se va a crear.</param>
            public RtfTreeNode(RtfNodeType nodeType)
            {
                type = nodeType;
                key = "";

                if (nodeType == RtfNodeType.Group || nodeType == RtfNodeType.Root)
                    children = new RtfNodeCollection();

                if (nodeType == RtfNodeType.Root)
                    root = this;
                
				/* Inicializados por defecto */
				//this.param = 0;
				//this.hasParam = false;
                //this.parent = null;
                //this.root = null;
                //this.tree = null;
                //children = null;
            }

            /// <summary>
            /// Constructor de la clase RtfTreeNode. Crea un nodo especificando su tipo, palabra clave y par醡etro.
            /// </summary>
            /// <param name="type">Tipo del nodo.</param>
            /// <param name="key">Palabra clave o s韒bolo de Control.</param>
            /// <param name="hasParameter">Indica si la palabra clave o el s韒bolo de Control va acompa馻do de un par醡etro.</param>
            /// <param name="parameter">Par醡etro del la palabra clave o s韒bolo de Control.</param>
            public RtfTreeNode(RtfNodeType type, string key, bool hasParameter, int parameter)
            {
                this.type = type;
                this.key = key;
                hasParam = hasParameter;
                param = parameter;

                if (type == RtfNodeType.Group || type == RtfNodeType.Root)
                    children = new RtfNodeCollection();

                if (type == RtfNodeType.Root)
                    root = this;

				/* Inicializados por defecto */
                //this.parent = null;
                //this.root = null;
                //this.tree = null;
                //children = null;
            }

            #endregion

            #region Constructor Privado

            /// <summary>
            /// Constructor privado de la clase RtfTreeNode. Crea un nodo a partir de un token del analizador l閤ico.
            /// </summary>
            /// <param name="token">Token RTF devuelto por el analizador l閤ico.</param>
            internal RtfTreeNode(RtfToken token)
            {
                type = (RtfNodeType)token.Type;
                key = token.Key;
                hasParam = token.HasParameter;
                param = token.Parameter;

				/* Inicializados por defecto */
                //this.parent = null;
                //this.root = null;
                //this.tree = null;
                //this.children = null;
            }

            #endregion

            #region M閠odos P鷅licos

            /// <summary>
            /// A馻de un nodo al final de la lista de hijos.
            /// </summary>
            /// <param name="newNode">Nuevo nodo a a馻dir.</param>
            public void AppendChild(RtfTreeNode newNode)
            {
				if(newNode != null)
				{
                    //Si a鷑 no ten韆 hijos se inicializa la colecci髇
                    if (children == null)
                        children = new RtfNodeCollection();

					//Se asigna como nodo padre el nodo actual
					newNode.parent = this;

                    //Se actualizan las propiedades Root y Tree del nuevo nodo y sus posibles hijos
                    updateNodeRoot(newNode);

					//Se a馻de el nuevo nodo al final de la lista de nodos hijo
					children.Add(newNode);
				}
            }

            /// <summary>
            /// Inserta un nuevo nodo en una posici髇 determinada de la lista de hijos.
            /// </summary>
            /// <param name="index">Posici髇 en la que se insertar?el nodo.</param>
            /// <param name="newNode">Nuevo nodo a insertar.</param>
            public void InsertChild(int index, RtfTreeNode newNode)
            {
                if (newNode != null)
                {
                    //Si a鷑 no ten韆 hijos se inicializa la colecci髇
                    if (children == null)
                        children = new RtfNodeCollection();

                    if (index >= 0 && index <= children.Count)
                    {
                        //Se asigna como nodo padre el nodo actual
                        newNode.parent = this;

                        //Se actualizan las propiedades Root y Tree del nuevo nodo y sus posibles hijos
                        updateNodeRoot(newNode);

                        //Se a馻de el nuevo nodo al final de la lista de nodos hijo
                        children.Insert(index, newNode);
                    }
                }
            }

            /// <summary>
            /// Elimina un nodo de la lista de hijos.
            /// </summary>
            /// <param name="index">Indice del nodo a eliminar.</param>
            public void RemoveChild(int index)
            {
                //Si el nodo actual tiene hijos
                if (children != null)
                {
                    if (index >= 0 && index < children.Count)
                    {
                        //Se elimina el i-閟imo hijo
                        children.RemoveAt(index);
                    }
                }
            }

            /// <summary>
            /// Elimina un nodo de la lista de hijos.
            /// </summary>
            /// <param name="node">Nodo a eliminar.</param>
            public void RemoveChild(RtfTreeNode node)
            {
                //Si el nodo actual tiene hijos
                if (children != null)
                {
                    //Se busca el nodo a eliminar
                    int index = children.IndexOf(node);
                    
                    //Si lo encontramos
                    if (index != -1)
                    {
                        //Se elimina el i-閟imo hijo
                        children.RemoveAt(index);
                    }
                }
            }

            /// <summary>
            /// Elimina un nodo de la lista de hijos.
            /// </summary>
            /// <param name="node">Nodo a eliminar.</param>
            public bool RemoveChildDeep(RtfTreeNode node)
            {
                //Si el nodo actual tiene hijos
                if (children != null)
                {
                    //Se busca el nodo a eliminar
                    int index = children.IndexOf(node);

                    //Si lo encontramos
                    if (index != -1)
                    {
                        //Se elimina el i-閟imo hijo
                        children.RemoveAt(index);
                        return true;
                    }
                    else
                    {
                        foreach (RtfTreeNode n in children)
                        {
                            if (n.RemoveChildDeep(node))
                            {
                                return true;
                            }
                        }

                        return false;
                    }
                }

                return false;
            }

            /// <summary>
            /// Elimina un nodo de la lista de hijos.
            /// </summary>
            /// <param name="node">Nodo a eliminar.</param>
            /// <param name="newnode">Nodo a eliminar.</param>
            public bool ReplaceChildDeep(RtfTreeNode node, RtfTreeNode newnode)
            {
                //Si el nodo actual tiene hijos
                if (children != null)
                {
                    //Se busca el nodo a eliminar
                    int index = children.IndexOf(node);

                    //Si lo encontramos
                    if (index != -1)
                    {
                        //Se elimina el i-閟imo hijo
                        children.RemoveAt(index);

                        InsertChild(index, newnode);
                        return true;
                    }
                    else
                    {
                        foreach (RtfTreeNode n in children)
                        {
                            if (n.ReplaceChildDeep(node, newnode))
                            {
                                return true;
                            }
                        }

                        return false;
                    }
                }

                return false;
            }

            /// <summary>
            /// Realiza una copia exacta del nodo actual.
            /// </summary>
            /// <returns>Devuelve una copia exacta del nodo actual.</returns>
            public RtfTreeNode CloneNode()
            {
                RtfTreeNode clon = new RtfTreeNode();

                clon.key = key;
                clon.hasParam = hasParam;
                clon.param = param;
                clon.parent = null;
                clon.root = null;
                clon.tree = null;
                clon.type = type;

                //Se clonan tambi閚 cada uno de los hijos
                clon.children = null;

                if (children != null)
                {
                    clon.children = new RtfNodeCollection();

                    foreach (RtfTreeNode child in children)
                    {
                        RtfTreeNode childclon = child.CloneNode();
                        childclon.parent = clon;

                        clon.children.Add(childclon);
                    }
                }

                return clon;
            }

            /// <summary>
            /// Indica si el nodo actual tiene nodos hijos.
            /// </summary>
            /// <returns>Devuelve true si el nodo actual tiene alg鷑 nodo hijo.</returns>
            public bool HasChildNodes()
            {
                bool res = false;

                if (children != null && children.Count > 0)
                    res = true;

                return res;
            }

            /// <summary>
            /// Devuelve el primer nodo de la lista de nodos hijos del nodo actual cuya palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <returns>Primer nodo de la lista de nodos hijos del nodo actual cuya palabra clave es la indicada como par醡etro.</returns>
            public RtfTreeNode SelectSingleChildNode(string keyword)
            {
                int i = 0;
                bool found = false;
                RtfTreeNode node = null;

                if (children != null)
                {
                    while (i < children.Count && !found)
                    {
                        if (children[i].key == keyword)
                        {
                            node = children[i];
                            found = true;
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Devuelve el primer nodo de la lista de nodos hijos del nodo actual cuyo tipo es el indicado como par醡etro.
            /// </summary>
            /// <param name="nodeType">Tipo de nodo buscado.</param>
            /// <returns>Primer nodo de la lista de nodos hijos del nodo actual cuyo tipo es el indicado como par醡etro.</returns>
            public RtfTreeNode SelectSingleChildNode(RtfNodeType nodeType)
            {
                int i = 0;
                bool found = false;
                RtfTreeNode node = null;

                if (children != null)
                {
                    while (i < children.Count && !found)
                    {
                        if (children[i].type == nodeType)
                        {
                            node = children[i];
                            found = true;
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Devuelve el primer nodo de la lista de nodos hijos del nodo actual cuya palabra clave y par醡etro son los indicados como par醡etros.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <param name="param">Par醡etro buscado.</param>
            /// <returns>Primer nodo de la lista de nodos hijos del nodo actual cuya palabra clave y par醡etro son los indicados como par醡etros.</returns>
            public RtfTreeNode SelectSingleChildNode(string keyword, int param)
            {
                int i = 0;
                bool found = false;
                RtfTreeNode node = null;

                if (children != null)
                {
                    while (i < children.Count && !found)
                    {
                        if (children[i].key == keyword && children[i].param == param)
                        {
                            node = children[i];
                            found = true;
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Devuelve el primer nodo grupo de la lista de nodos hijos del nodo actual cuya primera palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <returns>Primer nodo grupo de la lista de nodos hijos del nodo actual cuya primera palabra clave es la indicada como par醡etro.</returns>
            public RtfTreeNode SelectSingleChildGroup(string keyword)
            {
                return SelectSingleChildGroup(keyword, false);
            }

            /// <summary>
            /// Devuelve el primer nodo grupo de la lista de nodos hijos del nodo actual cuya primera palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <param name="ignoreSpecial">Si est?activo se ignorar醤 los nodos de control '\*' previos a algunas palabras clave.</param>
            /// <returns>Primer nodo grupo de la lista de nodos hijos del nodo actual cuya primera palabra clave es la indicada como par醡etro.</returns>
            public RtfTreeNode SelectSingleChildGroup(string keyword, bool ignoreSpecial)
            {
                int i = 0;
                bool found = false;
                RtfTreeNode node = null;

                if (children != null)
                {
                    while (i < children.Count && !found)
                    {
                        if (children[i].NodeType == RtfNodeType.Group && children[i].HasChildNodes() &&
                            (
                             (children[i].FirstChild.NodeKey == keyword) ||
                             (ignoreSpecial && children[i].ChildNodes[0].NodeKey == "*" && children[i].ChildNodes[1].NodeKey == keyword))
                            )
                        {
                            node = children[i];
                            found = true;
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Devuelve el primer nodo del 醨bol, a partir del nodo actual, cuyo tipo es el indicado como par醡etro.
            /// </summary>
            /// <param name="nodeType">Tipo del nodo buscado.</param>
            /// <returns>Primer nodo del 醨bol, a partir del nodo actual, cuyo tipo es el indicado como par醡etro.</returns>
            public RtfTreeNode SelectSingleNode(RtfNodeType nodeType)
            {
                int i = 0;
                bool found = false;
                RtfTreeNode node = null;

                if (children != null)
                {
                    while (i < children.Count && !found)
                    {
                        if (children[i].type == nodeType)
                        {
                            node = children[i];
                            found = true;
                        }
                        else
                        {
                            node = children[i].SelectSingleNode(nodeType);

                            if (node != null)
                            {
                                found = true;
                            }
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Devuelve el primer nodo del 醨bol, a partir del nodo actual, cuya palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <returns>Primer nodo del 醨bol, a partir del nodo actual, cuya palabra clave es la indicada como par醡etro.</returns>
            public RtfTreeNode SelectSingleNode(string keyword)
            {
                int i = 0;
                bool found = false;
                RtfTreeNode node = null;

                if (children != null)
                {
                    while (i < children.Count && !found)
                    {
                        if (children[i].key == keyword)
                        {
                            node = children[i];
                            found = true;
                        }
                        else
                        {
                            node = children[i].SelectSingleNode(keyword);

                            if (node != null)
                            {
                                found = true;
                            }
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Devuelve el primer nodo grupo del 醨bol, a partir del nodo actual, cuya primera palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <returns>Primer nodo grupo del 醨bol, a partir del nodo actual, cuya primera palabra clave es la indicada como par醡etro.</returns>
            public RtfTreeNode SelectSingleGroup(string keyword)
            {
                return SelectSingleGroup(keyword, false);
            }

            /// <summary>
            /// Devuelve el primer nodo grupo del 醨bol, a partir del nodo actual, cuya primera palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <param name="ignoreSpecial">Si est?activo se ignorar醤 los nodos de control '\*' previos a algunas palabras clave.</param>
            /// <returns>Primer nodo grupo del 醨bol, a partir del nodo actual, cuya primera palabra clave es la indicada como par醡etro.</returns>
            public RtfTreeNode SelectSingleGroup(string keyword, bool ignoreSpecial)
            {
                int i = 0;
                bool found = false;
                RtfTreeNode node = null;

                if (children != null)
                {
                    while (i < children.Count && !found)
                    {
                        if (children[i].NodeType == RtfNodeType.Group && children[i].HasChildNodes() &&
                            (
                             (children[i].FirstChild.NodeKey == keyword) ||
                             (ignoreSpecial && children[i].ChildNodes[0].NodeKey == "*" && children[i].ChildNodes[1].NodeKey == keyword))
                            )
                        {
                            node = children[i];
                            found = true;
                        }
                        else
                        {
                            node = children[i].SelectSingleGroup(keyword, ignoreSpecial);

                            if (node != null)
                            {
                                found = true;
                            }
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Devuelve el primer nodo del 醨bol, a partir del nodo actual, cuya palabra clave y par醡etro son los indicados como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <param name="param">Par醡etro buscado.</param>
            /// <returns>Primer nodo del 醨bol, a partir del nodo actual, cuya palabra clave y par醡etro son ls indicados como par醡etro.</returns>
            public RtfTreeNode SelectSingleNode(string keyword, int param)
            {
                int i = 0;
                bool found = false;
                RtfTreeNode node = null;

                if (children != null)
                {
                    while (i < children.Count && !found)
                    {
                        if (children[i].key == keyword && children[i].param == param)
                        {
                            node = children[i];
                            found = true;
                        }
                        else
                        {
                            node = children[i].SelectSingleNode(keyword, param);

                            if (node != null)
                            {
                                found = true;
                            }
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Devuelve todos los nodos, a partir del nodo actual, cuya palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <returns>Colecci髇 de nodos, a partir del nodo actual, cuya palabra clave es la indicada como par醡etro.</returns>
            public RtfNodeCollection SelectNodes(string keyword)
            {
                RtfNodeCollection nodes = new RtfNodeCollection();

                if (children != null)
                {
                    foreach (RtfTreeNode node in children)
                    {
                        if (node.key == keyword)
                        {
                            nodes.Add(node);
                        }

                        nodes.AddRange(node.SelectNodes(keyword));
                    }
                }

                return nodes;
            }

            /// <summary>
            /// Devuelve todos los nodos grupo, a partir del nodo actual, cuya primera palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <returns>Colecci髇 de nodos grupo, a partir del nodo actual, cuya primera palabra clave es la indicada como par醡etro.</returns>
            public RtfNodeCollection SelectGroups(string keyword)
            {
                return SelectGroups(keyword, false);
            }

            /// <summary>
            /// Devuelve todos los nodos grupo, a partir del nodo actual, cuya primera palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <param name="ignoreSpecial">Si est?activo se ignorar醤 los nodos de control '\*' previos a algunas palabras clave.</param>
            /// <returns>Colecci髇 de nodos grupo, a partir del nodo actual, cuya primera palabra clave es la indicada como par醡etro.</returns>
            public RtfNodeCollection SelectGroups(string keyword, bool ignoreSpecial)
            {
                RtfNodeCollection nodes = new RtfNodeCollection();

                if (children != null)
                {
                    foreach (RtfTreeNode node in children)
                    {
                        if (node.NodeType == RtfNodeType.Group && node.HasChildNodes() &&
                            (
                             (node.FirstChild.NodeKey == keyword) ||
                             (ignoreSpecial && node.ChildNodes[0].NodeKey == "*" && node.ChildNodes[1].NodeKey == keyword))
                            )
                        {
                            nodes.Add(node);
                        }

                        nodes.AddRange(node.SelectGroups(keyword, ignoreSpecial));
                    }
                }

                return nodes;
            }

            /// <summary>
            /// Devuelve todos los nodos, a partir del nodo actual, cuyo tipo es el indicado como par醡etro.
            /// </summary>
            /// <param name="nodeType">Tipo del nodo buscado.</param>
            /// <returns>Colecci髇 de nodos, a partir del nodo actual, cuyo tipo es la indicado como par醡etro.</returns>
            public RtfNodeCollection SelectNodes(RtfNodeType nodeType)
            {
                RtfNodeCollection nodes = new RtfNodeCollection();

                if (children != null)
                {
                    foreach (RtfTreeNode node in children)
                    {
                        if (node.type == nodeType)
                        {
                            nodes.Add(node);
                        }

                        nodes.AddRange(node.SelectNodes(nodeType));
                    }
                }

                return nodes;
            }

            /// <summary>
            /// Devuelve todos los nodos, a partir del nodo actual, cuya palabra clave y par醡etro son los indicados como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <param name="param">Par醡etro buscado.</param>
            /// <returns>Colecci髇 de nodos, a partir del nodo actual, cuya palabra clave y par醡etro son los indicados como par醡etro.</returns>
            public RtfNodeCollection SelectNodes(string keyword, int param)
            {
                RtfNodeCollection nodes = new RtfNodeCollection();

                if (children != null)
                {
                    foreach (RtfTreeNode node in children)
                    {
                        if (node.key == keyword && node.param == param)
                        {
                            nodes.Add(node);
                        }

                        nodes.AddRange(node.SelectNodes(keyword, param));
                    }
                }

                return nodes;
            }

            /// <summary>
            /// Devuelve todos los nodos de la lista de nodos hijos del nodo actual cuya palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <returns>Colecci髇 de nodos de la lista de nodos hijos del nodo actual cuya palabra clave es la indicada como par醡etro.</returns>
            public RtfNodeCollection SelectChildNodes(string keyword)
            {
                RtfNodeCollection nodes = new RtfNodeCollection();

                if (children != null)
                {
                    foreach (RtfTreeNode node in children)
                    {
                        if (node.key == keyword)
                        {
                            nodes.Add(node);
                        }
                    }
                }

                return nodes;
            }

            /// <summary>
            /// Devuelve todos los nodos grupos de la lista de nodos hijos del nodo actual cuya primera palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <returns>Colecci髇 de nodos grupo de la lista de nodos hijos del nodo actual cuya primera palabra clave es la indicada como par醡etro.</returns>
            public RtfNodeCollection SelectChildGroups(string keyword)
            {
                return SelectChildGroups(keyword, false);
            }

            /// <summary>
            /// Devuelve todos los nodos grupos de la lista de nodos hijos del nodo actual cuya primera palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <param name="ignoreSpecial">Si est?activo se ignorar醤 los nodos de control '\*' previos a algunas palabras clave.</param>
            /// <returns>Colecci髇 de nodos grupo de la lista de nodos hijos del nodo actual cuya primera palabra clave es la indicada como par醡etro.</returns>
            public RtfNodeCollection SelectChildGroups(string keyword, bool ignoreSpecial)
            {
                RtfNodeCollection nodes = new RtfNodeCollection();

                if (children != null)
                {
                    foreach (RtfTreeNode node in children)
                    {
                        if (node.NodeType == RtfNodeType.Group && node.HasChildNodes() &&
                            (
                             (node.FirstChild.NodeKey == keyword) ||
                             (ignoreSpecial && node.ChildNodes[0].NodeKey == "*" && node.ChildNodes[1].NodeKey == keyword))
                            )
                        {
                            nodes.Add(node);
                        }
                    }
                }

                return nodes;
            }

            /// <summary>
            /// Devuelve todos los nodos de la lista de nodos hijos del nodo actual cuyo tipo es el indicado como par醡etro.
            /// </summary>
            /// <param name="nodeType">Tipo del nodo buscado.</param>
            /// <returns>Colecci髇 de nodos de la lista de nodos hijos del nodo actual cuyo tipo es el indicado como par醡etro.</returns>
            public RtfNodeCollection SelectChildNodes(RtfNodeType nodeType)
            {
                RtfNodeCollection nodes = new RtfNodeCollection();

                if (children != null)
                {
                    foreach (RtfTreeNode node in children)
                    {
                        if (node.type == nodeType)
                        {
                            nodes.Add(node);
                        }
                    }
                }

                return nodes;
            }

            /// <summary>
            /// Devuelve todos los nodos de la lista de nodos hijos del nodo actual cuya palabra clave y par醡etro son los indicados como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <param name="param">Par醡etro buscado.</param>
            /// <returns>Colecci髇 de nodos de la lista de nodos hijos del nodo actual cuya palabra clave y par醡etro son los indicados como par醡etro.</returns>
            public RtfNodeCollection SelectChildNodes(string keyword, int param)
            {
                RtfNodeCollection nodes = new RtfNodeCollection();

                if (children != null)
                {
                    foreach (RtfTreeNode node in children)
                    {
                        if (node.key == keyword && node.param == param)
                        {
                            nodes.Add(node);
                        }
                    }
                }

                return nodes;
            }

            /// <summary>
            /// Devuelve el siguiente nodo hermano del actual cuya palabra clave es la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <returns>Primer nodo hermano del actual cuya palabra clave es la indicada como par醡etro.</returns>
            public RtfTreeNode SelectSibling(string keyword)
            {
                RtfTreeNode node = null;
                RtfTreeNode par = parent;

                if (par != null)
                {
                    int curInd = par.ChildNodes.IndexOf(this);

                    int i = curInd + 1;
                    bool found = false;

                    while (i < par.children.Count && !found)
                    {
                        if (par.children[i].key == keyword)
                        {
                            node = par.children[i];
                            found = true;
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Devuelve el siguiente nodo hermano del actual cuyo tipo es el indicado como par醡etro.
            /// </summary>
            /// <param name="nodeType">Tpo de nodo buscado.</param>
            /// <returns>Primer nodo hermano del actual cuyo tipo es el indicado como par醡etro.</returns>
            public RtfTreeNode SelectSibling(RtfNodeType nodeType)
            {
                RtfTreeNode node = null;
                RtfTreeNode par = parent;

                if (par != null)
                {
                    int curInd = par.ChildNodes.IndexOf(this);

                    int i = curInd + 1;
                    bool found = false;

                    while (i < par.children.Count && !found)
                    {
                        if (par.children[i].type == nodeType)
                        {
                            node = par.children[i];
                            found = true;
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Devuelve el siguiente nodo hermano del actual cuya palabra clave y par醡etro son los indicados como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <param name="param">Par醡etro buscado.</param>
            /// <returns>Primer nodo hermano del actual cuya palabra clave y par醡etro son los indicados como par醡etro.</returns>
            public RtfTreeNode SelectSibling(string keyword, int param)
            {
                RtfTreeNode node = null;
                RtfTreeNode par = parent;

                if (par != null)
                {
                    int curInd = par.ChildNodes.IndexOf(this);

                    int i = curInd + 1;
                    bool found = false;

                    while (i < par.children.Count && !found)
                    {
                        if (par.children[i].key == keyword && par.children[i].param == param)
                        {
                            node = par.children[i];
                            found = true;
                        }

                        i++;
                    }
                }

                return node;
            }

            /// <summary>
            /// Busca todos los nodos de tipo Texto que contengan el texto buscado.
            /// </summary>
            /// <param name="text">Texto buscado en el documento.</param>
            /// <returns>Lista de nodos, a partir del actual, que contienen el texto buscado.</returns>
            public RtfNodeCollection FindText(string text)
            {
                RtfNodeCollection list = new RtfNodeCollection();

                //Si el nodo actual tiene hijos
                if (children != null)
                {
                    foreach (RtfTreeNode node in children)
                    {
                        if (node.NodeType == RtfNodeType.Text && node.NodeKey.IndexOf(text) != -1)
                            list.Add(node);
                        else if(node.NodeType == RtfNodeType.Group)
                            list.AddRange(node.FindText(text));
                    }
                }

                return list;
            }

            /// <summary>
            /// Busca y reemplaza un texto determinado en todos los nodos de tipo Texto a partir del actual.
            /// </summary>
            /// <param name="oldValue">Texto buscado.</param>
            /// <param name="newValue">Texto de reemplazo.</param>
            public void ReplaceText(string oldValue, string newValue)
            {
                //Si el nodo actual tiene hijos
                if (children != null)
                {
                    foreach (RtfTreeNode node in children)
                    {
                        if (node.NodeType == RtfNodeType.Text)
                            node.NodeKey = node.NodeKey.Replace(oldValue, newValue);
                        else if (node.NodeType == RtfNodeType.Group)
                            node.ReplaceText(oldValue, newValue);
                    }
                }
            }

			/// <summary>
			/// Devuelve una representaci髇 del nodo donde se indica su tipo, clave, indicador de par醡etro y valor de par醡etro
			/// </summary>
			/// <returns>Cadena de caracteres del tipo [TIPO, CLAVE, IND_PARAMETRO, VAL_PARAMETRO]</returns>
			public override string ToString()
			{
				return "[" + type + ", " + key + ", " + hasParam + ", " + param + "]";
			}

            #endregion

            #region Metodos Privados

            /// <summary>
            /// Decodifica un caracter especial indicado por su c骴igo decimal
            /// </summary>
            /// <param name="code">C骴igo del caracter especial (\')</param>
            /// <param name="enc">Codificaci髇 utilizada para decodificar el caracter especial.</param>
            /// <returns>Caracter especial decodificado.</returns>
            private string DecodeControlChar(int code, Encoding enc)
            {
                //Contributed by Jan Stuchl韐
                return enc.GetString(new byte[] { (byte)code });
            }

            /// <summary>
            /// Obtiene el Texto RTF a partir de la representaci髇 en 醨bol del nodo actual.
            /// </summary>
            /// <returns>Texto RTF del nodo.</returns>
            private string getRtf()
            {
                string res = "";

                Encoding enc = tree.GetEncoding();

                res = getRtfInm(this, null, enc);

                return res;
            }

            /// <summary>
            /// M閠odo auxiliar para obtener el Texto RTF del nodo actual a partir de su representaci髇 en 醨bol.
            /// </summary>
            /// <param name="curNode">Nodo actual del 醨bol.</param>
            /// <param name="prevNode">Nodo anterior tratado.</param>
            /// <param name="enc">Codificaci髇 del documento.</param>
            /// <returns>Texto en formato RTF del nodo.</returns>
            private string getRtfInm(RtfTreeNode curNode, RtfTreeNode prevNode, Encoding enc)
            {
                StringBuilder res = new StringBuilder("");

                if (curNode.NodeType == RtfNodeType.Root)
                    res.Append("");
                else if (curNode.NodeType == RtfNodeType.Group)
                    res.Append("{");
                else
                {
                    if (curNode.NodeType == RtfNodeType.Control ||
                        curNode.NodeType == RtfNodeType.Keyword)
                    {
                        res.Append("\\");
                    }
                    else  //curNode.NodeType == RtfNodeType.Text
                    {
                        if (prevNode != null &&
                            prevNode.NodeType == RtfNodeType.Keyword)
                        {
                            int code = Char.ConvertToUtf32(curNode.NodeKey, 0);

                            if (code >= 32 && code < 128)
                                res.Append(" ");
                        }
                    }

                    AppendEncoded(res, curNode.NodeKey, enc);

                    if (curNode.HasParameter)
                    {
                        if (curNode.NodeType == RtfNodeType.Keyword)
                        {
                            res.Append(Convert.ToString(curNode.Parameter));
                        }
                        else if (curNode.NodeType == RtfNodeType.Control)
                        {
                            //Si es un caracter especial como las vocales acentuadas
                            if (curNode.NodeKey == "\'")
                            {						
								res.Append(GetHexa(curNode.Parameter));
                            }
                        }
                    }
                }

                //Se obtienen los nodos hijos
                RtfNodeCollection children = curNode.ChildNodes;

                //Si el nodo tiene hijos se obtiene el c骴igo RTF de los hijos
                if (children != null)
                {
                    for (int i = 0; i < children.Count; i++)
                    {
                        RtfTreeNode node = children[i];

                        if (i > 0)
                            res.Append(getRtfInm(node, children[i - 1], enc));
                        else
                            res.Append(getRtfInm(node, null, enc));
                    }
                }

                if (curNode.NodeType == RtfNodeType.Group)
                {
                    res.Append("}");
                }

                return res.ToString();
            }

            /// <summary>
            /// Concatena dos cadenas utilizando la codificaci髇 del documento.
            /// </summary>
            /// <param name="res">Cadena original.</param>
            /// <param name="s">Cadena a a馻dir.</param>
            /// <param name="enc">Codificaci髇 del documento.</param>
            private void AppendEncoded(StringBuilder res, string s, Encoding enc)
            {
                //Contributed by Jan Stuchl韐

                for (int i = 0; i < s.Length; i++)
                {
                    int code = Char.ConvertToUtf32(s, i);

                    if (code >= 128 || code < 32)
                    {
                        res.Append(@"\'");
                        byte[] bytes = enc.GetBytes(new char[] { s[i] });
                        res.Append(GetHexa(bytes[0]));
                    }
                    else
                    {
                        if ((s[i] == '{') || (s[i] == '}') || (s[i] == '\\'))
                        {
                            res.Append(@"\");
                        }

                        res.Append(s[i]);
                    }
                }
            }

            /// <summary>
            /// Obtiene el c骴igo hexadecimal de un entero.
            /// </summary>
            /// <param name="code">N鷐ero entero.</param>
            /// <returns>C骴igo hexadecimal del entero pasado como par醡etro.</returns>
            private string GetHexa(int code)
            {
                //Contributed by Jan Stuchl韐

                string hexa = Convert.ToString(code, 16);

                if (hexa.Length == 1)
                {
                    hexa = "0" + hexa;
                }

                return hexa;
            }

            /// <summary>
            /// Actualiza las propiedades Root y Tree de un nodo (y sus hijos) con las del nodo actual.
            /// </summary>
            /// <param name="node">Nodo a actualizar.</param>
            private void updateNodeRoot(RtfTreeNode node)
            {
                //Se asigna el nodo ra韟 del documento
                node.root = root;

                //Se asigna el 醨bol propietario del nodo
                node.tree = tree;

                //Si el nodo actualizado tiene hijos se actualizan tambi閚
                if (node.children != null)
                {
                    //Se actualizan recursivamente los hijos del nodo actual
                    foreach (RtfTreeNode nod in node.children)
                    {
                        updateNodeRoot(nod);
                    }
                }
            }

            /// <summary>
            /// Obtiene el texto contenido en el nodo actual.
            /// </summary>
            /// <param name="raw">Si este par醡etro est?activado se extraer?todo el texto contenido en el nodo, independientemente de si 閟te
            /// forma parte del texto real del documento.</param>
            /// <returns>Texto extraido del nodo.</returns>
            private string GetText(bool raw)
            {
                return GetText(raw, 1);
            }

            /// <summary>
            /// Obtiene el texto contenido en el nodo actual.
            /// </summary>
            /// <param name="raw">Si este par醡etro est?activado se extraer?todo el texto contenido en el nodo, independientemente de si 閟te
            /// forma parte del texto real del documento.</param>
            /// <param name="ignoreNchars">Ignore next N chars following \uN keyword</param>
            /// <returns>Texto extraido del nodo.</returns>
            private string GetText(bool raw, int ignoreNchars)
            {
                StringBuilder res = new StringBuilder("");

                if (NodeType == RtfNodeType.Group)
                {
                    int indkw = FirstChild.NodeKey.Equals("*") ? 1 : 0;

                    if (raw ||
                       (!ChildNodes[indkw].NodeKey.Equals("fonttbl") &&
                        !ChildNodes[indkw].NodeKey.Equals("colortbl") &&
                        !ChildNodes[indkw].NodeKey.Equals("stylesheet") &&
                        !ChildNodes[indkw].NodeKey.Equals("generator") &&
                        !ChildNodes[indkw].NodeKey.Equals("info") &&
                        !ChildNodes[indkw].NodeKey.Equals("pict") &&
                        !ChildNodes[indkw].NodeKey.Equals("object") &&
                        !ChildNodes[indkw].NodeKey.Equals("fldinst")))
                    {
                        if (ChildNodes != null)
                        {
                            int uc = ignoreNchars;
                            foreach (RtfTreeNode node in ChildNodes)
                            {
                                res.Append(node.GetText(raw, uc));

                                if (node.NodeType == RtfNodeType.Keyword && node.NodeKey.Equals("uc"))
                                    uc = node.Parameter;
                            }
                        }
                    }
                }
                else if (NodeType == RtfNodeType.Control)
                {
                    if (NodeKey == "'")
                        res.Append(DecodeControlChar(Parameter, tree.GetEncoding()));
                    else if (NodeKey == "~")  // non-breaking space
                        res.Append(" ");
                }
                else if (NodeType == RtfNodeType.Text)
                {
                    string newtext = NodeKey;

                    //Si el elemento anterior era un caracater Unicode (\uN) ignoramos los siguientes N caracteres
                    //seg鷑 la 鷏tima etiqueta \ucN
                    if (PreviousNode.NodeType == RtfNodeType.Keyword &&
                        PreviousNode.NodeKey.Equals("u"))
                    {
                        newtext = newtext.Substring(ignoreNchars);
                    }

                    res.Append(newtext);
                }
                else if (NodeType == RtfNodeType.Keyword)
                {
                    if (NodeKey.Equals("par"))
                        res.AppendLine("");
                    else if (NodeKey.Equals("tab"))
                        res.Append("\t");
                    else if (NodeKey.Equals("line"))
                        res.AppendLine("");
                    else if (NodeKey.Equals("lquote"))
                        res.Append("");
                    else if (NodeKey.Equals("rquote"))
                        res.Append("");
                    else if (NodeKey.Equals("ldblquote"))
                        res.Append("");
                    else if (NodeKey.Equals("rdblquote"))
                        res.Append("");
                    else if (NodeKey.Equals("emdash"))
                        res.Append("");
                    else if (NodeKey.Equals("u"))
                    {
                        res.Append(Char.ConvertFromUtf32(Parameter));
                    }
                }

                return res.ToString();
            }

            #endregion

            #region Propiedades

            /// <summary>
            /// Devuelve el nodo ra韟 del 醨bol del documento.
            /// </summary>
            /// <remarks>
            /// 蓅te no es el nodo ra韟 del 醨bol, sino que se trata simplemente de un nodo ficticio  de tipo ROOT del que parte el resto del 醨bol RTF.
            /// Tendr?por tanto un solo nodo hijo de tipo GROUP, raiz real del 醨bol.
			/// </remarks>
            public RtfTreeNode RootNode
            {
                get
                {
                    return root;
                }
				set
				{
					root = value;
				}
            }

            /// <summary>
            /// Devuelve el nodo padre del nodo actual.
            /// </summary>
            public RtfTreeNode ParentNode
            {
                get
                {
                    return parent;
                }
				set
				{
					parent = value;
				}
            }

            /// <summary>
            /// Devuelve el 醨bol Rtf al que pertenece el nodo.
            /// </summary>
            public RtfTree Tree
            {
                get
                {
                    return tree;
                }
                set
                {
                    tree = value;
                }
            }

            /// <summary>
            /// Devuelve el tipo del nodo actual.
            /// </summary>
            public RtfNodeType NodeType
            {
                get
                {
                    return type;
                }
				set
				{
					type = value;
				}
            }

            /// <summary>
            /// Devuelve la palabra clave, s韒bolo de Control o Texto del nodo actual.
            /// </summary>
            public string NodeKey
            {
                get
                {
                    return key;
                }
                set
                {
                    key = value;
                }
            }

            /// <summary>
            /// Indica si el nodo actual tiene par醡etro asignado.
            /// </summary>
            public bool HasParameter
            {
                get
                {
                    return hasParam;
                }
                set
                {
                    hasParam = value;
                }
            }

            /// <summary>
            /// Devuelve el par醡etro asignado al nodo actual.
            /// </summary>
            public int Parameter
            {
                get
                {
                    return param;
                }
                set
                {
                    param = value;
                }
            }

            /// <summary>
            /// Devuelve la colecci髇 de nodos hijo del nodo actual.
            /// </summary>
            public RtfNodeCollection ChildNodes
            {
                get
                {
                    return children;
                }
                set
                {
                    children = value;

                    foreach (RtfTreeNode node in children)
                    {
                        node.parent = this;

                        //Se actualizan las propiedades Root y Tree del nuevo nodo y sus posibles hijos
                        updateNodeRoot(node);
                    }
                }
            }

            /// <summary>
            /// Devuelve el primer nodo hijo cuya palabra clave sea la indicada como par醡etro.
            /// </summary>
            /// <param name="keyword">Palabra clave buscada.</param>
            /// <returns>Primer nodo hijo cuya palabra clave sea la indicada como par醡etro. En caso de no existir se devuelve null.</returns>
            public RtfTreeNode this[string keyword]
            {
                get
                {
                    return SelectSingleChildNode(keyword);
                }
            }

            /// <summary>
            /// Devuelve el hijo n-閟imo del nodo actual.
            /// </summary>
            /// <param name="childIndex">蚽dice del nodo hijo a recuperar.</param>
            /// <returns>Nodo hijo n-閟imo del nodo actual. Devuelve null en caso de no existir.</returns>
            public RtfTreeNode this[int childIndex]
            {
                get
                { 
                    RtfTreeNode res = null;

                    if (children != null && childIndex >= 0 && childIndex < children.Count)
                        res = children[childIndex];

                    return res;
                }
            }

            /// <summary>
            /// Devuelve el primer nodo hijo del nodo actual.
            /// </summary>
            public RtfTreeNode FirstChild
            {
                get
                {
                    RtfTreeNode res = null;

                    if (children != null && children.Count > 0)
                        res = children[0];

                    return res;
                }
            }

            /// <summary>
            /// Devuelve el 鷏timo nodo hijo del nodo actual.
            /// </summary>
            public RtfTreeNode LastChild
            {
                get
                {
                    RtfTreeNode res = null;

                    if (children != null && children.Count > 0)
                        return children[children.Count - 1];

                    return res;
                }
            }

            /// <summary>
            /// Devuelve el nodo hermano siguiente del nodo actual (Dos nodos son hermanos si tienen el mismo nodo padre [ParentNode]).
            /// </summary>
            public RtfTreeNode NextSibling
            {
                get
                {
                    RtfTreeNode res = null;

                    if (parent != null && parent.children != null)
                    {
                        int currentIndex = parent.children.IndexOf(this);

                        if (parent.children.Count > currentIndex + 1)
                            res = parent.children[currentIndex + 1];
                    }

                    return res;
                }
            }

            /// <summary>
            /// Devuelve el nodo hermano anterior del nodo actual (Dos nodos son hermanos si tienen el mismo nodo padre [ParentNode]).
            /// </summary>
            public RtfTreeNode PreviousSibling
            {
                get
                {
                    RtfTreeNode res = null;

                    if (parent != null && parent.children != null)
                    {
                        int currentIndex = parent.children.IndexOf(this);

                        if (currentIndex > 0)
                            res = parent.children[currentIndex - 1];
                    }

                    return res;
                }
            }

            /// <summary>
            /// Devuelve el siguiente nodo del 醨bol. 
            /// </summary>
            public RtfTreeNode NextNode
            {
                get
                {
                    RtfTreeNode res = null;

                    if (NodeType == RtfNodeType.Root)
                    {
                        res = FirstChild;
                    }
                    else if (parent != null && parent.children != null)
                    {
                        if (NodeType == RtfNodeType.Group && children.Count > 0)
                        {
                            res = FirstChild;
                        }
                        else
                        {
                            if (Index < (parent.children.Count - 1))
                            {
                                res = NextSibling;
                            }
                            else
                            {
                                res = parent.NextSibling;
                            }
                        }
                    }

                    return res;
                }
            }

            /// <summary>
            /// Devuelve el nodo anterior del 醨bol.
            /// </summary>
            public RtfTreeNode PreviousNode
            {
                get
                {
                    RtfTreeNode res = null;

                    if (NodeType == RtfNodeType.Root)
                    {
                        res = null;
                    }
                    else if (parent != null && parent.children != null)
                    {
                        if (Index > 0)
                        {
                            if (PreviousSibling.NodeType == RtfNodeType.Group)
                            {
                                res = PreviousSibling.LastChild;
                            }
                            else
                            {
                                res = PreviousSibling;
                            }
                        }
                        else
                        {
                            res = parent;
                        }
                    }

                    return res;
                }
            }

            /// <summary>
            /// Devuelve el c骴igo RTF del nodo actual y todos sus nodos hijos.
            /// </summary>
            public string Rtf
            {
                get
                {
                    return getRtf();
                }
            }

            /// <summary>
            /// Devuelve el 韓dice del nodo actual dentro de la lista de hijos de su nodo padre.
            /// </summary>
            public int Index
            {
                get
                {
                    int res = -1;

                    if(parent != null)
                        res = parent.children.IndexOf(this);

                    return res;
                }
            }

            /// <summary>
            /// Devuelve el fragmento de texto del documento contenido en el nodo actual.
            /// </summary>
            public string Text
            {
                get
                {
                    return GetText(false);
                }
            }

            /// <summary>
            /// Devuelve todo el texto contenido en el nodo actual.
            /// </summary>
            public string RawText
            {
                get
                {
                    return GetText(true);
                }
            }

            #endregion
        }
    }
}
