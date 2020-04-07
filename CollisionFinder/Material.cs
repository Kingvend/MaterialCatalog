namespace CollisionFinder
{
    /// <summary>
    /// Материалы МТР
    /// </summary>
    class Material
    {
        /// <summary>
        /// Код материала
        /// </summary>
        public string MaterialCode { get; set; }
        /// <summary>
        /// краткое наименование материала
        /// </summary>
        public string MaterialName { get; set; }
        /// <summary>
        /// Полное наименование материала
        /// </summary>
        public string MaterialFullName { get; set; }
        /// <summary>
        /// Единица измерения
        /// </summary>
        public string MaterialMeasureUnit { get; set; }
        /// <summary>
        /// Количество к закупу
        /// </summary>
        public string MaterialCountMU { get; set; }
        /// <summary>
        /// Порядковый номер строки в исходном файле
        /// </summary>
        public string MaterialRowNumber { get; set; }
        /// <summary>
        /// имя исходного файла, где содержится запись
        /// </summary>
        public string MaterialSource { get; set; }


    }
}



