namespace PowerXML.Tests;

public class PXmlWorkBookTests
{
    [Test]
    public void CreateFileTest()
    {
        var fileData = new PXmlWorkBook().CreateFile();
        
        Assert.That(fileData.Data.Length, Is.GreaterThan(0));
    }
}