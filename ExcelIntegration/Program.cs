
Console.WriteLine("Loading...");
var xlcsxIntegrator = new XLSX_Integrator();

List<Carro> carros = new List<Carro> 
{
    new Carro {Marca = "Audi", Modelo = "A1", Ano = 1999},
    new Carro {Marca = "Chevrolet", Modelo = "Astra", Ano = 2002},
    new Carro {Marca = "Fiat", Modelo = "Uno", Ano = 1990},
    new Carro {Marca = "Fiat", Modelo = "Palio", Ano = 2010},
};

xlcsxIntegrator.CreateFile(carros);
xlcsxIntegrator.ReadFile();