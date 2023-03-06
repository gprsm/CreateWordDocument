namespace CreateWordDocument.Models
{
    public class PersonTypeNum
    {
        public enum PersonType
        {            
            Family = 1,            
            Colleague = 2,
        }
        public enum Gender  
        {
            Man=1,
            Woman=2,
            Religious=3,
            Children=4
        }
        public enum ColumnType
        {
            Name,
            Family,
            Gender,
            PersonType,
            Company,
            Score,
            Signature,
            Text
        }
    }
}