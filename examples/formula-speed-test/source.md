# 公式转换速度测试源文本

## OMML 公式区

行内公式一：$E=mc^2$。

行内公式二：$a_i+b_i=c_i$。

行内公式三：$\lim_{x\to0}\frac{\sin x}{x}=1$。

独立公式一：

$$\int_0^1 x^2\,dx=\frac{1}{3}$$

独立公式二：

$$\begin{pmatrix}1&2\\3&4\end{pmatrix}\begin{pmatrix}x\\y\end{pmatrix}=\begin{pmatrix}5\\6\end{pmatrix}$$

独立公式三：

$$\sum_{k=1}^{n}k=\frac{n(n+1)}{2}$$

## Word 内 LaTeX 文本区

行内公式四：$f(x)=\frac{1}{\sqrt{2\pi}\sigma}e^{-\frac{(x-\mu)^2}{2\sigma^2}}$。

行内公式五：$z=\frac{x-\mu}{\sigma}$。

行内公式六：$\nabla\cdot\mathbf{E}=\frac{\rho}{\varepsilon_0}$。

独立公式四：

$$\int_{-\infty}^{\infty}e^{-x^2}\,dx=\sqrt{\pi}$$

独立公式五：

$$f(x)=\begin{cases}x^2,&x\ge0\\-x,&x<0\end{cases}$$

独立公式六：

$$\frac{\partial u}{\partial t}=\alpha\nabla^2u+f(x,t)$$

> 说明：`input.docx` 中前 6 个公式为 OMML，后 6 个公式为 Word 内
> LaTeX 文本。这是为了同时验证 OMML 与 LaTeX 两条转换路径。
