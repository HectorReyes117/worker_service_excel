namespace WorkerServiceExcel.Entities
{
    internal class ParameterWrite<T>
    {
        public int Row { get; set; }
        public int Col { get; set; }
        public string? Path { get; set; }
        public T? Value { get; set; }
    }
}
