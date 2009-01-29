using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace MACROAPI_dot_net_test_harness
{
    public class ParameterInfo
    {
        #region constructors
        public ParameterInfo()
        {
            _paramName = "";
            _dbType = DbType.Object;
            _paramDirection = ParameterDirection.Input;
            _value = null;
        }

        public ParameterInfo(string paramName, DbType dbType, ParameterDirection paramDirection, object value)
        {
            _paramName = paramName;
            _dbType = dbType;
            _paramDirection = paramDirection;
            _value = value;
        }
        #endregion

        #region properties
        private string _paramName;
        private DbType _dbType;
        private ParameterDirection _paramDirection;
        private object _value;

        /// <summary>
        /// Parameter Name 
        /// </summary>
        public string ParameterName
        {
            get
            {
                return _paramName;
            }
            set
            {
                _paramName = value;
            }
        }

        /// <summary>
        /// Generic DBType
        /// </summary>
        public DbType DbGenericType
        {
            get
            {
                return _dbType;
            }
            set
            {
                _dbType = value;
            }
        }

        /// <summary>
        /// Parameter Direction
        /// </summary>
        public ParameterDirection ParamDirection
        {
            get
            {
                return _paramDirection;
            }
            set
            {
                _paramDirection = value;
            }
        }

        /// <summary>
        /// Parameter value
        /// </summary>
        public object ParamValue
        {
            get
            {
                return _value;
            }
            set
            {
                _value = value;
            }
        }
        #endregion
    }
}
