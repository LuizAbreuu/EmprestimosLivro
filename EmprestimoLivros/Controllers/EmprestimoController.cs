using ClosedXML.Excel;
using EmprestimoLivros.Data;
using EmprestimoLivros.Models;
using EmprestimoLivros.Services.EmprestimosService;
using EmprestimoLivros.Services.SessaoService;
using Microsoft.AspNetCore.Mvc;
using System.Data;

namespace EmprestimoLivros.Controllers
{
    public class EmprestimoController : Controller
    {
        readonly private ISessaoInterface _sessaoInterface;
        private readonly IEmprestimosInterface _emprestimosInterface;

        public EmprestimoController(IEmprestimosInterface emprestimosInterface , ISessaoInterface sessaoInterface)
        {
            _sessaoInterface = sessaoInterface;
            _emprestimosInterface = emprestimosInterface;
        }

        public async Task<IActionResult> Index()
        {
            var usuario = _sessaoInterface.BuscarSessao();
            if (usuario == null) 
            {
                return RedirectToAction("Login", "Login");
            }

            var emprestimos = await _emprestimosInterface.BuscarEmprestimos();

            return View(emprestimos.Dados);
        }

        public IActionResult Cadastrar() 
        {
            var usuario = _sessaoInterface.BuscarSessao();
            if (usuario == null)
            {
                return RedirectToAction("Login", "Login");
            }
            return View();
        }

        [HttpGet]
        public async Task<IActionResult> Editar(int? id)
        {
            var usuario = _sessaoInterface.BuscarSessao();
            if (usuario == null)
            {
                return RedirectToAction("Login", "Login");
            }

            var emprestimo = await _emprestimosInterface.BuscarEmprestimosPorId(id);

            return View(emprestimo.Dados);
        }

        [HttpPost]
        public async Task<IActionResult> Cadastrar(EmprestimosModel emprestimos)
        {
            if (ModelState.IsValid) 
            {
                
                var emprestimoResult = await _emprestimosInterface.CadastrarEmprestimo(emprestimos);

                if (emprestimoResult.Status) 
                {
                    TempData["MensagemSucesso"] = emprestimoResult.Mensagem;
                }
                else
                {
                    TempData["MensagemErro"] = emprestimoResult.Mensagem;
                    return View(emprestimos);
                }
                
                return RedirectToAction("Index");
            }

            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Editar(EmprestimosModel emprestimos)
        {
            if (ModelState.IsValid)
            {
                var emprestimoResult = await _emprestimosInterface.EditarEmprestimo(emprestimos);

                if (emprestimoResult.Status) 
                {
                    TempData["MensagemSucesso"] = emprestimoResult.Mensagem;
                }
                else
                {
                    TempData["MensagemErro"] = emprestimoResult.Mensagem;
                    return View(emprestimos);
                }

                return RedirectToAction("Index");
            }
            TempData["MensagemErro"] = "Ocorreu algum erro no momento da edição";
            return View(emprestimos);
        }

        [HttpGet]
        public async Task<IActionResult> Excluir(int? id) 
        {
            var usuario = _sessaoInterface.BuscarSessao();
            if (usuario == null)
            {
                return RedirectToAction("Login", "Login");
            }

            var emprestimo = await _emprestimosInterface.BuscarEmprestimosPorId(id);

            return View(emprestimo.Dados);
        }

        public async Task<IActionResult> Exportar()
        {
            var dados = await _emprestimosInterface.BuscarDadosEmprestimoExcel();

            using (XLWorkbook workbook = new XLWorkbook()) 
            {
                workbook.AddWorksheet(dados, "Dados Empréstimo");

                using (MemoryStream ms = new MemoryStream()) 
                {
                    workbook.SaveAs(ms);

                    return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spredsheetml.sheet", "Emprestimo.xls");
                }
            }
        }

        [HttpPost]
        public async Task<IActionResult> Excluir(EmprestimosModel emprestimos) 
        {
            if(emprestimos == null)
            {
                TempData["MensagemErro"] = "Empréstimo não localizado";
                return View(emprestimos); 
            }

            var emprestimoResult = await _emprestimosInterface.RemoveEmprestimo(emprestimos);

            if (emprestimoResult.Status)
            {
                TempData["MensagemSucesso"] = emprestimoResult.Mensagem;
            }
            else
            {
                TempData["MensagemErro"] = emprestimoResult.Mensagem;
                return View(emprestimos);
            }

            return RedirectToAction("Index");
        }
    }
}
