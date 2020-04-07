namespace CollisionFinder
{
    /// <summary>
    /// класс содержащий информацию о файле с МТР, необходимую для составления отчета
    /// </summary>
    class MaterialFile
    {
        /// <summary>
        /// путь к файлу
        /// </summary>
        public string FilePath;
        /// <summary>
        /// порядковый номер последней строки
        /// </summary>
        public int LastRow;
        /// <summary>
        /// порядковый номер коонк с кодом материалов
        /// </summary>
        public int CodeCol;
        /// <summary>
        /// порядковый номер колонки с кратким наименованием материалов
        /// </summary>
        public int NameCol;
        /// <summary>
        /// порядковый номер колонки с полным наименованием материалов (1 часть)
        /// </summary>
        public int FullNameCol_1;
        /// <summary>
        /// порядковый номер колонки с полным наименованием материалов (2 часть)
        /// </summary>
        public int FullNameCol_2;
        /// <summary>
        /// порядковый номер колонки с единицей измерения 
        /// </summary>
        public int MessureCol;
        /// <summary>
        /// порядковый номер колонки с числом закупок 
        /// </summary>
        public int CountMesCol;       
    }
}
