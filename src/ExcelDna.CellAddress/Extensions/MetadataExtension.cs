using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Security;
using System.Text;

namespace ExcelDna.Extensions {
    /// <summary>
    ///     元数据扩展方法
    /// </summary>
    public static class MetadataExtension {
        /// <summary>
        ///     返回 满足条件<see cref="condition" />的 Linq 集合 索引
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        public static int IndexOf<T>(this IEnumerable<T> list, Predicate<T> condition) {
            if (list == null) {
                throw new ArgumentNullException(nameof(list));
            }
            if (condition == null) {
                throw new ArgumentNullException(nameof(condition));
            }

            int index = -1;
            foreach (T item in list) {
                index++;
                if (condition(item)) {
                    return index;
                }
            }
            return -1;
        }


        /// <summary>
        ///     检查表达式类型,默认类型为 <see cref="ExpressionType.MemberAccess" />
        /// </summary>
        /// <param name="expression"></param>
        /// <param name="expressionType"></param>
        private static void AssertExpression(this LambdaExpression expression,
                                             ExpressionType expressionType) {
            if (expression.Body.NodeType != expressionType) {
                throw new MemberAccessException($"{expression.Body.NodeType}类型的表达式无效,{expression}");
            }
        }
        #region MemberAccess
        /// <summary>
        ///     设置成员值
        ///     调用方 假设 属性表达式为 成员属性 并且必须具备写方法 <see cref="PropertyInfo.CanWrite" />
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="t"></param>
        /// <param name="expression"></param>
        /// <param name="value"></param>
        public static void SetMemberValue<TModel, TValue>(this TModel t,
                                                          Expression<Func<TModel, TValue>> expression,
                                                          TValue value) {
            if (expression == null) {
                throw new ArgumentNullException(nameof(expression));
            }
            MemberInfo member = expression.Body.GetMemberInfo();
            var propertyInfo = member as PropertyInfo;
            if (propertyInfo != null) {
                var property = propertyInfo;
                property.SetValue(t, value, new object[] { });
                return;
            }
            var fieldInfo = member as FieldInfo;
            if (fieldInfo != null) {
                var field = fieldInfo;
                field.SetValue(t, value);
                return;
            }
            throw new MemberAccessException($"属性 {member.Name} 不支持的成员类型 {member.GetType()}");
        }

        /// <summary>
        ///     获取成员值
        ///     调用方 假设 属性表达式为 成员属性或字段 并且必须具备读方法 <see cref="PropertyInfo.CanRead" />
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="t"></param>
        /// <param name="expression"></param>
        public static TValue GetMemberValue<TModel, TValue>(this TModel t, Expression<Func<TModel, TValue>> expression) {
            if (expression == null) {
                throw new ArgumentNullException(nameof(expression));
            }
            MemberInfo member = expression.Body.GetMemberInfo();
            if (member == null) {
                throw new MemberAccessException($"{expression.Body.NodeType}类型的表达式无效,{expression}");
            }
            if (member is PropertyInfo) {
                var property = ((PropertyInfo)member);
                return (TValue)property.GetValue(t, new object[] { });
            }
            if (member is FieldInfo) {
                var property = ((FieldInfo)member);
                return (TValue)property.GetValue(t);
            }
            throw new MemberAccessException($"属性 {member.Name} 不支持的成员类型 {member.GetType()}");
        }

        #endregion MemberAccess

        /// <summary>
        ///     根据 <see cref="DefaultValueAttribute" /> 获取对象的默认属性
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="t"></param>
        /// <param name="expression"></param>
        /// <returns></returns>
        public static TValue GetDefaultValue<TModel, TValue>(this TModel t, Expression<Func<TModel, TValue>> expression)
            where TModel : class {
            if (t == null) {
                throw new ArgumentNullException(nameof(t));
            }
            if (expression == null) {
                throw new ArgumentNullException(nameof(expression));
            }

            // Gets the attributes for the property.
            expression.AssertExpression(ExpressionType.MemberAccess);
            AttributeCollection attributes = TypeDescriptor.GetProperties(t)[MemberName(expression)].Attributes;
            var attribute = (DefaultValueAttribute)attributes[typeof(DefaultValueAttribute)];
            return (TValue)attribute.Value;
        }

        /// <summary>
        ///     在指定 String 数组的每个元素之间串联指定的分隔符 String，从而产生单个串联的字符串
        ///     参见 <see cref="String.Join(string,string[])" />
        /// </summary>
        /// <param name="separator"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public static string Join(this string separator, IEnumerable<string> values) {
            if (values == null) {
                throw new ArgumentNullException(nameof(values));
            }
            if (separator == null) {
                separator = string.Empty;
            }
            using (IEnumerator<string> enumerator = values.GetEnumerator()) {
                if (!enumerator.MoveNext()) {
                    return string.Empty;
                }
                var builder = new StringBuilder();
                if (enumerator.Current != null) {
                    builder.Append(enumerator.Current);
                }
                while (enumerator.MoveNext()) {
                    builder.Append(separator);
                    if (enumerator.Current != null) {
                        builder.Append(enumerator.Current);
                    }
                }
                return builder.ToString();
            }
        }

        #region LoadType

        /// <summary>
        ///     根据名称加载类型
        /// </summary>
        /// <param name="typeName"></param>
        /// <returns></returns>
        public static Type LoadType(this string typeName) {
            if (string.IsNullOrEmpty(typeName)) {
                return null;
            }
            Type itemType = GetTypeFromString(typeName, false, false);
            if (itemType != null) {
                return itemType;
            }
            return null;
        }

        /// <summary>
        ///     根据 程序集名称 加载程序集
        /// </summary>
        /// <param name="assemblyName"></param>
        /// <returns></returns>
        public static Assembly LoadAssembly(this string assemblyName) {
            if (!string.IsNullOrEmpty(assemblyName)) {
                try {
                    return Assembly.Load(assemblyName);
                } catch (Exception ex) {
                    Trace.WriteLine("加载程序集 " + assemblyName + " 发生错误," + ex.Message);
                }
            }
            return null;
        }

        /// <summary>
        ///     根据类型名称获得类型对象
        /// </summary>
        /// <param name="typeName"></param>
        /// <param name="throwOnError"></param>
        /// <param name="ignoreCase"></param>
        /// <returns></returns>
        private static Type GetTypeFromString(string typeName, bool throwOnError, bool ignoreCase) {
            return GetTypeFromString(Assembly.GetCallingAssembly(), typeName, throwOnError, ignoreCase);
        }

        public static Type GetTypeFromString(Assembly relativeAssembly,
                                             string typeName,
                                             bool throwOnError,
                                             bool ignoreCase) {
            // Check if the type name specifies the assembly name
            if (typeName.IndexOf(',') == -1) {
                // Attempt to lookup the type from the relativeAssembly
                Type type = relativeAssembly.GetType(typeName, false, ignoreCase);
                if (type != null) {
                    // Found type in relative assembly
                    return type;
                }
                Assembly[] loadedAssemblies = null;
                try {
                    loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();
                } catch (SecurityException) {
                    // Insufficient permissions to get the list of loaded assemblies
                }
                if (loadedAssemblies != null) {
                    // Search the loaded assemblies for the type
                    foreach (Assembly assembly in loadedAssemblies) {
                        type = assembly.GetType(typeName, false, ignoreCase);
                        if (type != null) {
                            // Found type in loaded assembly
                            return type;
                        }
                    }
                }
                // Didn't find the type
                if (throwOnError) {
                    throw new TypeLoadException("不能加载类型 [" + typeName + "]. Tried assembly [" +
                                                relativeAssembly.FullName +
                                                "] and all loaded assemblies");
                }
                return null;
            }
            // Includes explicit assembly name
            //LogLog.Debug("SystemInfo: Loading type ["+typeName+"] from global Type");
            return Type.GetType(typeName, throwOnError, ignoreCase);
        }

        #endregion LoadType

        #region DisplayName

        /// <summary>
        ///     获取类型 modelType 的 <see cref="DisplayNameAttribute">显示名属性内容</see>
        /// </summary>
        /// <param name="modelType"></param>
        /// <returns></returns>
        public static string DisplayName(this Type modelType) {
            DisplayNameAttribute attributes =
                TypeDescriptor.GetAttributes(modelType).OfType<DisplayNameAttribute>().FirstOrDefault();
            return attributes != null ? attributes.DisplayName : modelType.Name;
        }

        /// <summary>
        ///     根据表达式获取 成员显示名称
        ///     <seealso cref="DisplayNameAttribute" />
        ///     对于枚举类型 使用 DescriptionAttribute 属性
        ///     <seealso cref="DescriptionAttribute" />
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <returns></returns>
        public static string DisplayName<TModel>(string memberName) {
            PropertyDescriptor propertyInfo = TypeDescriptor.GetProperties(typeof(TModel))[memberName];
            if (propertyInfo != null) {
                return propertyInfo.DisplayName;
            }
            return memberName;
        }

        /// <summary>
        ///     根据表达式获取 成员显示名称
        ///     <seealso cref="DisplayNameAttribute" />
        ///     对于枚举类型 使用 DescriptionAttribute 属性
        ///     <seealso cref="DescriptionAttribute" />
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="expression"></param>
        /// <returns></returns>
        public static string DisplayName<TModel, TValue>(this Expression<Func<TModel, TValue>> expression) {
            switch (expression.Body.NodeType) {
                case ExpressionType.MemberAccess: {
                        var body = (MemberExpression)expression.Body;
                        var propertyInfo = body.Member as PropertyInfo;
                        if (propertyInfo != null) {
                            PropertyDescriptor propertyDesc =
                                TypeDescriptor.GetProperties(typeof(TModel)).Find(propertyInfo.Name, false);
                            return propertyDesc.DisplayName;
                        } else {
                            return body.Member.Name;
                        }
                    }
            }
            throw new MemberAccessException($"{expression.Body.NodeType}类型的表达式无效,{expression}");
        }

        #endregion DisplayName

        #region MemberType

        /// <summary>
        ///     根据表达式获取 成员名称
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="expression"></param>
        /// <returns></returns>
        public static Type MemberType<TModel, TValue>(this Expression<Func<TModel, TValue>> expression)
            where TModel : class {
            if (expression == null) {
                throw new ArgumentNullException(nameof(expression));
            }
            var memberInfo = expression.GetMemberInfo();
            var propertyInfo = memberInfo as PropertyInfo;
            if (propertyInfo != null) {
                return propertyInfo.PropertyType;
            }
            var fieldInfo = memberInfo as FieldInfo;
            if (fieldInfo != null) {
                return fieldInfo.FieldType;
            }
            var methodInfo = memberInfo as MethodInfo;
            if (methodInfo != null) {
                return methodInfo.ReturnType;
            }
            throw new NotSupportedException("不支持的成员信息类型," + memberInfo.GetType());
        }

        #endregion MemberType

        #region MemberName

        /// <summary>
        ///     根据表达式获取 成员名称
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="expression"></param>
        /// <returns></returns>
        public static string MemberName<TModel>(this Expression<Func<TModel, object>> expression) {
            return expression.GetMemberInfo().Name;
        }

        /// <summary>
        ///     根据表达式获取 成员名称
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="expression"></param>
        /// <returns></returns>
        public static string MemberName<TModel, TValue>(this Expression<Func<TModel, TValue>> expression)
            where TModel : class {
            if (expression == null) {
                throw new ArgumentNullException(nameof(expression));
            }
            return expression.GetMemberInfo().Name;
        }

        /// <summary>
        ///     根据表达式获取 成员名称
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="t"></param>
        /// <param name="expression"></param>
        /// <returns></returns>
        public static string MemberName<TModel, TValue>(this TModel t, Expression<Func<TModel, TValue>> expression)
            where TModel : class {
            if (t == null) {
                throw new ArgumentNullException(nameof(t));
            }
            return expression.GetMemberInfo().Name;
        }

        /// <summary>
        ///     根据表达式获取 成员名称
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="t"></param>
        /// <param name="expression"></param>
        /// <returns></returns>
        public static string MemberName<TModel, TValue>(this IEnumerable<TModel> t,
                                                        Expression<Func<TModel, TValue>> expression)
            where TModel : class {
            if (t == null) {
                throw new ArgumentNullException(nameof(t));
            }
            return expression.GetMemberInfo().Name;
        }

        #endregion MemberName

        #region DefaultValue

        /// <summary>
        ///     <see cref="Nullable{T}">可空结构</see> 默认值
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static T Default<T>(this T? value) where T : struct {
            return value ?? new T();
        }

        /// <summary>
        ///     <see cref="Nullable{T}">可空结构</see> 默认值
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue">若为空时给定的默认值</param>
        /// <returns></returns>
        public static T Default<T>(this T? value, T defaultValue) where T : struct {
            return value ?? defaultValue;
        }

        #endregion DefaultValue

        #region GetMemberInfo 获取 表达式的 成员信息

        /// <summary>
        /// 获取 表达式中的成员信息
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="?"></param>
        /// <param name="model"></param>
        /// <param name="expression"></param>
        /// <returns></returns>
        public static MemberInfo GetMemberInfo<TModel, TValue>(this TModel model, Expression<Func<TModel, TValue>> expression) {
            if (expression == null) {
                throw new ArgumentNullException(nameof(expression));
            }
            return expression.GetMemberInfo();
        }

        /// <summary>
        ///     获取 表达式的 成员信息
        ///     如 r=> r.Name
        /// </summary>
        /// <param name="expression"></param>
        /// <returns></returns>
        private static MemberInfo GetMemberInfo(this Expression expression) {
            switch (expression.NodeType) {
                case ExpressionType.Lambda:
                    return GetMemberInfo(expression as LambdaExpression);
                case ExpressionType.MemberAccess:
                    return GetMemberInfo(expression as MemberExpression);
                case ExpressionType.Convert:
                    return GetMemberInfo(expression as UnaryExpression);
                case ExpressionType.Call:
                    return GetMemberInfo(expression as MethodCallExpression);
                default:
                    throw new NotSupportedException($"不支持的表达式类型:{expression.NodeType}\n表达式:{expression}");
            }
        }
        /// <summary>
        /// 获取 <see cref="LambdaExpression"/>表达式的 成员信息
        /// </summary>
        /// <param name="lambda"></param>
        /// <returns></returns>
        private static MemberInfo GetMemberInfo(this LambdaExpression lambda) {
            return GetMemberInfo(lambda.Body);
        }
        /// <summary>
        /// 获取 <see cref="UnaryExpression"/>表达式的 成员信息
        /// </summary>
        /// <param name="unary"></param>
        /// <returns></returns>
        private static MemberInfo GetMemberInfo(this UnaryExpression unary) {
            return GetMemberInfo(unary.Operand);
        }
        /// <summary>
        /// 获取 <see cref="MemberExpression"/>表达式的 成员信息
        /// </summary>
        /// <param name="member"></param>
        /// <returns></returns>
        private static MemberInfo GetMemberInfo(this MemberExpression member) {
            return member.Member;
        }
        /// <summary>
        /// 获取 <see cref="MethodCallExpression"/>表达式的 成员信息
        /// </summary>
        /// <param name="lambda"></param>
        /// <returns></returns>
        private static MemberInfo GetMemberInfo(this MethodCallExpression lambda) {
            return lambda.Method;
        }
        #endregion GetMemberInfo
    }
}